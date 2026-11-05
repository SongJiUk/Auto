using DocumentFormat.OpenXml.Packaging;
using FormFiller.Core.Config;
using FormFiller.Core.Models;

namespace FormFiller.Core.Excel;

/// <summary>
/// config.json 의 DocumentConfig 를 읽어 범용으로 엑셀 파일을 생성하는 엔진.
///
/// [전체 흐름]
/// 1. LoadPreviousDay : 어제 파일 → PrevDayCount / PrevCumulativeCount 채우기
/// 2. Generate        : DocumentConfig 에 따라 엑셀 파일 생성
///
/// [원칙]
/// - samples/ 원본 템플릿은 절대 수정하지 않는다 (복사본에만 쓴다).
/// - 수식 셀(누계, 총누계, 합계)은 절대 덮어쓰지 않는다.
/// - 이미지·도형·병합·서식·열 너비는 건드리지 않는다.
/// </summary>
public static class GenericExcelWriter
{
    // 한국어 요일 배열 — DayOfWeek 순서(0=일요일)와 일치
    private static readonly string[] KoreanDays = ["일", "월", "화", "수", "목", "금", "토"];

    // ── 공개 API ──────────────────────────────────────────────────────

    /// <summary>
    /// 어제 파일에서 equipment_table 의 누계(todayCol+2번째 열)와
    /// personnel_table 의 총누계(todayCol+2번째 열)를 읽어
    /// DailyInput.Equipment[i].PrevDayCount 와 Personnel[i].PrevCumulativeCount 에 채운다.
    ///
    /// [누계 열 계산 방법]
    /// - equipment_table: todayCol(H) → 열인덱스(8) + 2 → 열인덱스(10) → "J"
    /// - personnel_table: todayCol(P) → 열인덱스(16) + 2 → 열인덱스(18) → "R"
    /// 수식 셀이므로 config 에 명시하지 않고 여기서 자동 계산한다.
    /// </summary>
    /// <param name="prevFilePath">어제 공사일보_출력일보 파일 경로</param>
    /// <param name="input">오늘 입력 데이터 (PrevDay 필드가 채워진다)</param>
    /// <param name="dailyDoc">공사일보 DocumentConfig</param>
    public static void LoadPreviousDay(string prevFilePath, DailyInput input, DocumentConfig dailyDoc)
    {
        if (!File.Exists(prevFilePath)) return;

        using var doc   = SpreadsheetDocument.Open(prevFilePath, isEditable: false);
        var       wbPart = doc.WorkbookPart!;
        var       wsPart = CellHelper.FindWorksheetPart(wbPart, dailyDoc.SheetName);
        if (wsPart == null) return;

        // ── equipment_table 필드 찾기 ────────────────────────────────
        var eqField = dailyDoc.Fields.FirstOrDefault(f => f.Type == FieldTypes.EquipmentTable);
        if (eqField != null && !string.IsNullOrEmpty(eqField.NameCol) && !string.IsNullOrEmpty(eqField.TodayCol))
        {
            // 누계 열 = todayCol(H) 의 열인덱스 + 2 (수식 셀이므로 직접 쓰지 않고 읽기만 한다)
            string cumCol = IndexToColLetter(CellHelper.ColLetterToIndex(eqField.TodayCol) + 2);

            for (int row = eqField.StartRow; row <= eqField.EndRow; row++)
            {
                string nameAddr = $"{eqField.NameCol}{row}";
                string cumAddr  = $"{cumCol}{row}";

                string? cellName = CellHelper.GetCellValue(wsPart, nameAddr);
                if (string.IsNullOrWhiteSpace(cellName)) continue;

                var entry = input.Equipment
                                 .FirstOrDefault(e => NormName(e.Name) == NormName(cellName));
                if (entry == null) continue;

                entry.PrevDayCount = CellHelper.GetCellInt(wsPart, cumAddr) ?? 0;
            }
        }

        // ── personnel_table 필드 찾기 ────────────────────────────────
        var pnField = dailyDoc.Fields.FirstOrDefault(f => f.Type == FieldTypes.PersonnelTable);
        if (pnField != null && !string.IsNullOrEmpty(pnField.RoleCol) && !string.IsNullOrEmpty(pnField.TodayCol))
        {
            // 총누계 열 = todayCol(P) 의 열인덱스 + 2
            string cumCol = IndexToColLetter(CellHelper.ColLetterToIndex(pnField.TodayCol) + 2);

            for (int row = pnField.StartRow; row <= pnField.EndRow; row++)
            {
                string roleAddr = $"{pnField.RoleCol}{row}";
                string cumAddr  = $"{cumCol}{row}";

                string? cellRole = CellHelper.GetCellValue(wsPart, roleAddr);
                if (string.IsNullOrWhiteSpace(cellRole)) continue;

                var entry = input.Personnel
                                 .FirstOrDefault(p => NormName(p.Role) == NormName(cellRole));
                if (entry == null) continue;

                entry.PrevCumulativeCount = CellHelper.GetCellInt(wsPart, cumAddr) ?? 0;
            }
        }
    }

    /// <summary>
    /// DocumentConfig 에 따라 엑셀 파일을 생성하고, 생성된 파일의 전체 경로를 반환한다.
    ///
    /// [처리 흐름]
    /// 1. templateFile → output 복사 (원본 보존)
    /// 2. SpreadsheetDocument 열기
    /// 3. fields 를 순회하며 타입별로 셀에 값 쓰기
    /// 4. SetFullCalcOnLoad → 수식 재계산 플래그 설정
    /// 5. 저장 후 경로 반환
    ///
    /// docConfig.IncludesSujangjae 가 true 이면 personnel_table 에 수장재 인원까지 포함한다.
    /// tbm_personnel 타입은 항상 수장재 포함 합계(TotalPersonnelWithSujangjae)를 사용한다.
    /// </summary>
    /// <param name="input">오늘 입력 데이터</param>
    /// <param name="docConfig">처리할 문서 설정</param>
    /// <param name="baseDir">exe 위치 (상대경로 기준점)</param>
    /// <param name="outputFolder">출력 폴더명 (baseDir 기준 상대경로)</param>
    public static string Generate(DailyInput input, DocumentConfig docConfig,
                                  string baseDir, string outputFolder)
    {
        // ── 1. 경로 계산 ─────────────────────────────────────────────
        string templatePath = Path.Combine(baseDir, docConfig.TemplateFile);
        if (!File.Exists(templatePath))
            throw new FileNotFoundException(
                $"템플릿 파일을 찾을 수 없습니다.\n경로: {templatePath}");

        string outputPath = BuildOutputPath(input.Date, docConfig.OutputPattern, baseDir, outputFolder);

        // ── 2. 원본 복사 — samples/ 는 절대 수정하지 않는다 ──────────
        EnsureDirectoryExists(outputPath);
        File.Copy(templatePath, outputPath, overwrite: true);

        // ── 3. 복사본을 편집 모드로 열기 ─────────────────────────────
        using var doc    = SpreadsheetDocument.Open(outputPath, isEditable: true);
        var       wbPart = doc.WorkbookPart!;
        var       wsPart = CellHelper.FindWorksheetPart(wbPart, docConfig.SheetName)
                           ?? throw new InvalidOperationException(
                                  $"시트 '{docConfig.SheetName}'를 찾을 수 없습니다.\n파일: {outputPath}");

        // ── 4. fields 순회: 타입별 처리 ──────────────────────────────
        foreach (var field in docConfig.Fields)
            WriteField(field, input, docConfig, doc, wsPart);

        // ── 5. 수식 재계산 플래그 설정 (J열·R열·합계행 수식 보장) ────
        CellHelper.SetFullCalcOnLoad(wbPart);
        wbPart.Workbook.Save();

        return outputPath;
    }

    // ── 타입별 처리 ───────────────────────────────────────────────────

    /// <summary>
    /// 단일 FieldConfig 를 타입에 따라 처리한다.
    /// docConfig.IncludesSujangjae 로 인원 테이블 범위를 결정한다.
    /// </summary>
    private static void WriteField(FieldConfig field, DailyInput input,
                                   DocumentConfig docConfig,
                                   SpreadsheetDocument doc, WorksheetPart wsPart)
    {
        switch (field.Type)
        {
            // ── 날짜: "2026년 11월 05일 목요일" ──────────────────────
            case FieldTypes.DateKorean:
                if (!string.IsNullOrEmpty(field.Cell))
                    CellHelper.SetCellString(doc, wsPart, field.Cell,
                                             FormatKoreanDate(input.Date));
                break;

            // ── 날짜: "2026.11.05" ────────────────────────────────────
            case FieldTypes.DateSimple:
                if (!string.IsNullOrEmpty(field.Cell))
                    CellHelper.SetCellString(doc, wsPart, field.Cell,
                                             input.Date.ToString("yyyy.MM.dd"));
                break;

            // ── 텍스트 ───────────────────────────────────────────────
            case FieldTypes.Text:
                if (!string.IsNullOrEmpty(field.Cell))
                    CellHelper.SetCellString(doc, wsPart, field.Cell,
                                             GetTextValue(field.Id, input));
                break;

            // ── 멀티라인 텍스트 → 셀 범위에 나눠 쓰기 ───────────────
            case FieldTypes.Multiline:
                if (field.Cells is { Count: > 0 })
                    WriteMultiline(doc, wsPart, GetMultilineValue(field.Id, input), field.Cells);
                break;

            // ── 장비 투입현황 테이블 ──────────────────────────────────
            case FieldTypes.EquipmentTable:
                WriteEquipmentTable(wsPart, input.Equipment, field);
                break;

            // ── 인원 투입현황 테이블 ─────────────────────────────────
            // docConfig.IncludesSujangjae = true 이면 수장재 포함 목록 사용
            case FieldTypes.PersonnelTable:
                WritePersonnelTable(wsPart, input.GetPersonnelFor(docConfig.IncludesSujangjae), field);
                break;

            // ── 금일 대수 > 0 장비명 쉼표 연결 ──────────────────────
            case FieldTypes.ActiveEquipmentNames:
                if (!string.IsNullOrEmpty(field.Cell))
                {
                    string names = string.Join(", ",
                        input.Equipment.Where(e => e.TodayCount > 0).Select(e => e.Name));
                    CellHelper.SetCellString(doc, wsPart, field.Cell, names);
                }
                break;

            // ── TBM 인원 — 수장재 포함 합계 사용 ────────────────────
            case FieldTypes.TbmPersonnel:
                if (!string.IsNullOrEmpty(field.Cell))
                    CellHelper.SetCellInt(wsPart, field.Cell, input.TotalPersonnelWithSujangjae);
                break;

            // ── 참석율 — 계획 = 실시이므로 항상 100 ─────────────────
            case FieldTypes.AttendanceRate:
                if (!string.IsNullOrEmpty(field.Cell))
                    CellHelper.SetCellInt(wsPart, field.Cell, 100);
                break;
        }
    }

    /// <summary>
    /// 멀티라인 텍스트를 셀 범위에 나눠 쓴다.
    /// 텍스트 줄 수가 셀 범위보다 적으면 나머지 셀은 빈 문자열로 클리어한다.
    /// </summary>
    private static void WriteMultiline(SpreadsheetDocument doc, WorksheetPart wsPart,
                                       string text, List<string> cells)
    {
        string[] lines = text.Split(["\r\n", "\n"], StringSplitOptions.None);

        for (int i = 0; i < cells.Count; i++)
        {
            string value = i < lines.Length ? lines[i] : "";
            CellHelper.SetCellString(doc, wsPart, cells[i], value);
        }
    }

    /// <summary>
    /// 장비 투입현황 테이블에 PrevDayCount 와 TodayCount 를 쓴다.
    /// NameCol 로 행을 찾고, 없는 행은 건너뛴다.
    /// 누계 열(TodayCol+2)은 수식이므로 절대 쓰지 않는다.
    /// </summary>
    private static void WriteEquipmentTable(WorksheetPart wsPart,
                                            List<EquipmentEntry> equipment,
                                            FieldConfig field)
    {
        if (string.IsNullOrEmpty(field.NameCol) ||
            string.IsNullOrEmpty(field.PrevDayCol) ||
            string.IsNullOrEmpty(field.TodayCol))
            return;

        for (int row = field.StartRow; row <= field.EndRow; row++)
        {
            string nameAddr = $"{field.NameCol}{row}";
            string? cellName = CellHelper.GetCellValue(wsPart, nameAddr);
            if (string.IsNullOrWhiteSpace(cellName)) continue;

            var entry = equipment.FirstOrDefault(e => NormName(e.Name) == NormName(cellName));
            if (entry == null) continue;

            // F열 = 전일(대), H열 = 금일(대) — 누계(J열)는 수식이므로 쓰지 않는다
            CellHelper.SetCellInt(wsPart, $"{field.PrevDayCol}{row}", entry.PrevDayCount);
            CellHelper.SetCellInt(wsPart, $"{field.TodayCol}{row}",   entry.TodayCount);
        }
    }

    /// <summary>
    /// 인원 투입현황 테이블에 PrevCumulativeCount 와 TodayCount 를 쓴다.
    /// RoleCol 로 행을 찾는다.
    /// 총누계 열(TodayCol+2)은 수식이므로 절대 쓰지 않는다.
    /// </summary>
    private static void WritePersonnelTable(WorksheetPart wsPart,
                                            List<PersonnelEntry> personnel,
                                            FieldConfig field)
    {
        if (string.IsNullOrEmpty(field.RoleCol) ||
            string.IsNullOrEmpty(field.PrevCumulativeCol) ||
            string.IsNullOrEmpty(field.TodayCol))
            return;

        for (int row = field.StartRow; row <= field.EndRow; row++)
        {
            string roleAddr = $"{field.RoleCol}{row}";
            string? cellRole = CellHelper.GetCellValue(wsPart, roleAddr);
            if (string.IsNullOrWhiteSpace(cellRole)) continue;

            var entry = personnel.FirstOrDefault(p => NormName(p.Role) == NormName(cellRole));
            if (entry == null) continue;

            // N열 = 전일누계(명), P열 = 금일(명) — 총누계(R열)는 수식이므로 쓰지 않는다
            CellHelper.SetCellInt(wsPart, $"{field.PrevCumulativeCol}{row}", entry.PrevCumulativeCount);
            CellHelper.SetCellInt(wsPart, $"{field.TodayCol}{row}",          entry.TodayCount);
        }
    }

    // ── 값 조회 헬퍼 ─────────────────────────────────────────────────

    /// <summary>
    /// text 타입 필드의 Id 로 DailyInput 에서 값을 가져온다.
    /// </summary>
    private static string GetTextValue(string fieldId, DailyInput input) => fieldId switch
    {
        "weather"       => input.Weather,
        "today_work"    => input.TodayWork,
        "tomorrow_work" => input.TomorrowWork,
        _               => "",
    };

    /// <summary>
    /// multiline 타입 필드의 Id 로 DailyInput 에서 멀티라인 텍스트를 가져온다.
    /// </summary>
    private static string GetMultilineValue(string fieldId, DailyInput input) => fieldId switch
    {
        "today_work"    => input.TodayWork,
        "tomorrow_work" => input.TomorrowWork,
        _               => "",
    };

    // ── 경로·날짜 헬퍼 ───────────────────────────────────────────────

    /// <summary>
    /// OutputPattern 의 {yy}, {MM}, {dd} 를 날짜로 치환해 출력 경로를 반환한다.
    /// 예) "공사일보_출력일보_{yy}.{MM}.{dd}.xlsx" → ".../output/공사일보_출력일보_26.11.05.xlsx"
    /// </summary>
    private static string BuildOutputPath(DateTime date, string outputPattern,
                                          string baseDir, string outputFolder)
    {
        string fileName = ApplyDatePattern(outputPattern, date);
        string dir      = Path.Combine(baseDir, outputFolder);
        return Path.Combine(dir, fileName);
    }

    /// <summary>
    /// 패턴의 {yy}, {MM}, {dd} 를 날짜 값으로 치환한다.
    /// </summary>
    public static string ApplyDatePattern(string pattern, DateTime date) =>
        pattern.Replace("{yy}", date.ToString("yy"))
               .Replace("{MM}", date.ToString("MM"))
               .Replace("{dd}", date.ToString("dd"));

    /// <summary>출력 폴더가 없으면 만든다.</summary>
    private static void EnsureDirectoryExists(string filePath)
    {
        string? dir = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);
    }

    /// <summary>
    /// 날짜를 한국어 요일 포함 형식으로 변환한다.
    /// 예) 2026-11-05 → "2026년 11월 05일 목요일"
    /// </summary>
    private static string FormatKoreanDate(DateTime date)
    {
        string dayOfWeek = KoreanDays[(int)date.DayOfWeek];
        return $"{date:yyyy년 MM월 dd일} {dayOfWeek}요일";
    }

    /// <summary>
    /// 1-based 열 인덱스 → 열 문자열로 변환한다.
    /// 예) 10 → "J",  18 → "R",  27 → "AA"
    ///
    /// [계산 원리]
    /// 1-based 를 0-based 로 변환한 뒤 26진법으로 변환한다.
    /// </summary>
    private static string IndexToColLetter(uint index)
    {
        string result = "";
        while (index > 0)
        {
            index--;                             // 0-based 로 조정
            result = (char)('A' + index % 26) + result;
            index /= 26;
        }
        return result;
    }

    /// <summary>
    /// 셀 텍스트 비교용 정규화: 앞뒤 공백 제거, 대소문자 소문자 통일.
    /// "B/H 380 " vs "B/H 380" 처럼 공백 차이를 무시하기 위해 사용.
    /// </summary>
    private static string NormName(string s)
        => s.Trim().ToLowerInvariant();
}
