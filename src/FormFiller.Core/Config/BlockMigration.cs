namespace FormFiller.Core.Config;

/// <summary>
/// 예전 방식 항목(날씨·금일 작업·장비 표처럼 뜻이 코드에 정해진 종류)을 블록으로 바꾼다.
/// 설정 화면의 [블록 방식으로 바꾸기] 에서 쓴다. 바꾼 뒤에도 같은 칸에 같은 값이 들어가도록 한다(테스트로 확인).
/// 사진·그림 위치·시트별 입력은 그대로 둔다.
/// </summary>
public static class BlockMigration
{
    // 예전 방식에서 쓰던 값 이름. 여러 문서가 이 이름으로 같은 입력을 함께 쓰게 된다.
    public const string Weather = "날씨", TodayWork = "금일 작업", TomorrowWork = "명일 작업",
                        Equipment = "장비", Personnel = "인원", Sujangjae = "수장재 인원";

    /// <summary>예전 방식 항목이 있는지</summary>
    public static bool HasLegacy(DocumentConfig doc) => doc.Fields.Any(f => IsLegacy(f.Type));

    public static bool IsLegacy(string type) => type is FieldTypes.DateKorean or FieldTypes.DateSimple or FieldTypes.Text
        or FieldTypes.Number or FieldTypes.Multiline or FieldTypes.EquipmentTable or FieldTypes.PersonnelTable
        or FieldTypes.ActiveEquipmentNames or FieldTypes.TbmPersonnel or FieldTypes.AttendanceRate;

    /// <summary>문서의 예전 항목을 블록으로 바꾼 새 항목 목록 (원본은 바꾸지 않는다)</summary>
    public static List<FieldConfig> Convert(DocumentConfig doc)
    {
        var result = new List<FieldConfig>();
        bool needsSujang = false;

        foreach (var f in doc.Fields)
        {
            if (!IsLegacy(f.Type)) { result.Add(f); continue; }

            var b = new FieldConfig { Id = f.Id, Label = f.Label, Cell = f.Cell, Placeholder = f.Placeholder, Suffix = f.Suffix };
            switch (f.Type)
            {
                case FieldTypes.DateKorean:
                    b.Type = BlockTypes.Date; b.DateFormat = "{yyyy}년 {MM}월 {dd}일 {요일}"; break;
                case FieldTypes.DateSimple:
                    b.Type = BlockTypes.Date; b.DateFormat = "{yyyy}.{MM}.{dd}"; break;

                case FieldTypes.Text:
                case FieldTypes.Number:
                    b.Type = BlockTypes.Value;
                    b.Key  = f.Id switch { "weather" => Weather, "today_work" => TodayWork, "tomorrow_work" => TomorrowWork, _ => null };
                    break;

                case FieldTypes.Multiline:
                    b.Type   = BlockTypes.Lines;
                    b.Cells  = f.Cells is null ? null : [.. f.Cells];
                    // 앞 글자는 붙이지 않는다: 예전 화면은 사용자가 입력한 "- " 를 그대로 썼다 (같은 결과 보장)
                    (b.Key, b.StartFrom) = f.Id switch
                    {
                        "today_work"    => (TodayWork, TomorrowWork),   // 금일 = 어제의 명일 (없으면 어제 금일)
                        "tomorrow_work" => (TomorrowWork, "same"),
                        _               => ((string?)null, "same"),
                    };
                    break;

                case FieldTypes.EquipmentTable:
                    b.Type = BlockTypes.Table; b.Key = Equipment;
                    b.StartRow = f.StartRow; b.EndRow = f.EndRow;
                    b.Columns = Cols((BlockTypes.RoleName, f.NameCol), (BlockTypes.RoleSpec, f.SpecCol),
                                     (BlockTypes.RolePrev, f.PrevDayCol), (BlockTypes.RoleToday, f.TodayCol),
                                     (BlockTypes.RoleTotal, Plus2(f.TodayCol)));
                    break;

                case FieldTypes.PersonnelTable:
                    b.Type = BlockTypes.Table; b.Key = Personnel;
                    b.StartRow = f.StartRow; b.EndRow = f.EndRow;
                    b.Columns = Cols((BlockTypes.RoleName, f.RoleCol), (BlockTypes.RolePrev, f.PrevCumulativeCol),
                                     (BlockTypes.RoleToday, f.TodayCol), (BlockTypes.RoleTotal, Plus2(f.TodayCol)));
                    break;

                case FieldTypes.ActiveEquipmentNames:
                    b.Type = BlockTypes.Calc; b.CalcKind = BlockTypes.CalcNames; b.Sources = [Equipment]; break;

                case FieldTypes.TbmPersonnel:
                    // 예전: 공사 인원 + 수장재 인원 합계
                    b.Type = BlockTypes.Calc; b.CalcKind = BlockTypes.CalcSum; b.Sources = [Personnel, Sujangjae];
                    needsSujang = true;
                    break;

                case FieldTypes.AttendanceRate:
                    b.Type = BlockTypes.Value; b.ValueSource = BlockTypes.SourceFixed; b.FixedValue = "100"; break;
            }
            result.Add(b);
        }

        // 수장재 인원: 칸은 없고 계산에만 쓰는 매일 입력 숫자
        if (needsSujang && !result.Any(f => BlockTypes.KeyOf(f) == Sujangjae))
            result.Add(new FieldConfig { Id = "sujangjae", Label = Sujangjae, Type = BlockTypes.Value });

        return result;
    }

    private static List<ColumnRole> Cols(params (string Role, string? Col)[] cols) =>
        cols.Where(c => !string.IsNullOrEmpty(c.Col)).Select(c => new ColumnRole { Role = c.Role, Col = c.Col! }).ToList();

    /// <summary>예전 엔진은 누계 열을 "금일 열 + 2" 로 정했다 (예: H → J)</summary>
    private static string? Plus2(string? col)
    {
        if (string.IsNullOrEmpty(col)) return null;
        uint idx = Excel.CellHelper.ColLetterToIndex(col) + 2;
        string r = "";
        while (idx > 0) { idx--; r = (char)('A' + idx % 26) + r; idx /= 26; }
        return r;
    }
}
