using FormFiller.Core.Config;

namespace FormFiller.App.ViewModels.Settings;

/// <summary>
/// 설정 화면에서 한 필드(FieldConfig)를 편집하는 행 모델.
///
/// 필드 타입에 따라 세 가지 UI 그룹이 있다:
///   - 단순 셀 (IsSimpleType=true) : 항목명 | 셀 주소
///   - 멀티라인 (IsMultiline=true) : 항목명 | 시작 셀 주소
///   - 테이블   (IsTableType=true) : 항목명 | 시작행·끝행·열 설정
/// </summary>
public class FieldEditRow : ViewModelBase
{
    // ── 불변 정보 ────────────────────────────────────────────────────

    public string FieldId { get; }
    public string Label   { get; }
    public string Type    { get; }

    // ── 타입 판별 (Visibility 바인딩용) ──────────────────────────────

    public bool IsTableType   => Type is FieldTypes.EquipmentTable or FieldTypes.PersonnelTable;
    public bool IsEquipTable  => Type == FieldTypes.EquipmentTable;
    public bool IsPersonTable => Type == FieldTypes.PersonnelTable;
    public bool IsSimpleType  => !IsTableType;

    // ── 단순 셀 / 멀티라인 시작 셀 주소 ─────────────────────────────

    private string _cellAddress = "";

    /// <summary>
    /// 단순 셀이면 해당 셀 주소 (예: "C6"),
    /// 멀티라인이면 시작 셀 주소 (예: "A9").
    /// </summary>
    public string CellAddress
    {
        get => _cellAddress;
        set => SetField(ref _cellAddress, value);
    }

    // ── 테이블 공통 ──────────────────────────────────────────────────

    private int _startRow;
    public int StartRow { get => _startRow; set => SetField(ref _startRow, value); }

    private int _endRow;
    public int EndRow { get => _endRow; set => SetField(ref _endRow, value); }

    private string _todayCol = "";
    public string TodayCol { get => _todayCol; set => SetField(ref _todayCol, value); }

    // ── equipment_table 전용 ─────────────────────────────────────────

    private string _nameCol = "";
    public string NameCol { get => _nameCol; set => SetField(ref _nameCol, value); }

    private string _prevDayCol = "";
    public string PrevDayCol { get => _prevDayCol; set => SetField(ref _prevDayCol, value); }

    // ── personnel_table 전용 ─────────────────────────────────────────

    private string _roleCol = "";
    public string RoleCol { get => _roleCol; set => SetField(ref _roleCol, value); }

    private string _prevCumulativeCol = "";
    public string PrevCumulativeCol { get => _prevCumulativeCol; set => SetField(ref _prevCumulativeCol, value); }

    // ── 생성자 ──────────────────────────────────────────────────────

    public FieldEditRow(FieldConfig f)
    {
        FieldId = f.Id;
        Label   = f.Label;
        Type    = f.Type;

        // 멀티라인은 Cells의 첫 번째 주소를 시작 셀로 사용
        _cellAddress = (f.Type == FieldTypes.Multiline)
            ? f.Cells?.FirstOrDefault() ?? ""
            : f.Cell ?? "";

        _startRow          = f.StartRow;
        _endRow            = f.EndRow;
        _nameCol           = f.NameCol           ?? "";
        _prevDayCol        = f.PrevDayCol        ?? "";
        _todayCol          = f.TodayCol          ?? "";
        _roleCol           = f.RoleCol           ?? "";
        _prevCumulativeCol = f.PrevCumulativeCol ?? "";
    }

    // ── 변환 ─────────────────────────────────────────────────────────

    /// <summary>현재 편집 상태를 FieldConfig로 변환한다.</summary>
    public FieldConfig ToFieldConfig()
    {
        if (IsTableType)
        {
            var fc = new FieldConfig
            {
                Id       = FieldId,
                Label    = Label,
                Type     = Type,
                StartRow = StartRow,
                EndRow   = EndRow,
            };

            if (IsEquipTable)
            {
                fc.NameCol    = NameCol;
                fc.PrevDayCol = PrevDayCol;
                fc.TodayCol   = TodayCol;
            }
            else
            {
                fc.RoleCol           = RoleCol;
                fc.PrevCumulativeCol = PrevCumulativeCol;
                fc.TodayCol          = TodayCol;
            }
            return fc;
        }

        if (Type == FieldTypes.Multiline)
        {
            return new FieldConfig
            {
                Id    = FieldId,
                Label = Label,
                Type  = Type,
                Cells = BuildCellRange(_cellAddress, 9),
            };
        }

        return new FieldConfig
        {
            Id    = FieldId,
            Label = Label,
            Type  = Type,
            Cell  = string.IsNullOrWhiteSpace(_cellAddress) ? null : _cellAddress,
        };
    }

    private static List<string> BuildCellRange(string startCell, int count)
    {
        if (string.IsNullOrWhiteSpace(startCell)) return [];

        int i = 0;
        while (i < startCell.Length && char.IsLetter(startCell[i])) i++;

        string col = startCell[..i].ToUpperInvariant();
        if (!int.TryParse(startCell[i..], out int startRow)) return [];

        return Enumerable.Range(startRow, count).Select(r => $"{col}{r}").ToList();
    }
}
