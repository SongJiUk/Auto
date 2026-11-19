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
    /// <param name="prevFilePath">이전 날짜 공사일보_출력일보 파일 경로 (FindPreviousFile 로 찾는다)</param>
    /// <param name="input">오늘 입력 데이터 (PrevDay 필드가 채워진다)</param>
    /// <param name="dailyDoc">공사일보 DocumentConfig</param>
    /// <remarks>
    /// 누계 칸은 수식이라, 프로그램이 만든 파일을 Excel 에서 열어 저장하기 전까지는
    /// 저장된 값(캐시)이 서식 원본 숫자 그대로다. 그래서 "전일 + 금일" 값 칸(프로그램이 직접 쓴 칸)을
    /// 더해서 누계를 구한다. 전일 열 설정이 없을 때만 수식 칸 값을 읽는다.
    /// </remarks>
    public static void LoadPreviousDay(string prevFilePath, DailyInput input, DocumentConfig dailyDoc)
        => LoadPreviousDayCore(prevFilePath, input, dailyDoc);

    /// <summary>
    /// 파일(앞 날짜 파일 또는 서식)의 장비·인원 표에 적힌 줄 목록을 읽는다.
    /// 화면의 장비·인원 줄을 이것으로 만들어서, 파일에 있는 줄(예: B/H 6W, 세륜기)이 빠지지 않게 한다.
    /// </summary>
    public static (List<(string Name, string Spec)> Equipment, List<string> Roles) ReadTableRows(
        string filePath, DocumentConfig dailyDoc)
    {
        var eq = new List<(string, string)>();
        var roles = new List<string>();
        if (!File.Exists(filePath)) return (eq, roles);

        using var doc  = CellHelper.OpenReadOnly(filePath);
        var       ws   = CellHelper.FindWorksheetPart(doc.WorkbookPart!, dailyDoc.SheetName);
        if (ws == null) return (eq, roles);

        var eqField = dailyDoc.Fields.FirstOrDefault(f => f.Type == FieldTypes.EquipmentTable);
        if (eqField is { NameCol.Length: > 0 })
            for (int row = eqField.StartRow; row <= eqField.EndRow; row++)
            {
                string? name = CellHelper.GetCellValue(ws, $"{eqField.NameCol}{row}")?.Trim();
                if (string.IsNullOrEmpty(name)) continue;
                string spec = string.IsNullOrEmpty(eqField.SpecCol) ? ""
                            : CellHelper.GetCellValue(ws, $"{eqField.SpecCol}{row}")?.Trim() ?? "";
                eq.Add((name, spec));
            }

        var pnField = dailyDoc.Fields.FirstOrDefault(f => f.Type == FieldTypes.PersonnelTable);
        if (pnField is { RoleCol.Length: > 0 })
            for (int row = pnField.StartRow; row <= pnField.EndRow; row++)
            {
                string? role = CellHelper.GetCellValue(ws, $"{pnField.RoleCol}{row}")?.Trim();
                if (!string.IsNullOrEmpty(role)) roles.Add(role);
            }

        return (eq, roles);
    }

    private static void LoadPreviousDayCore(string prevFilePath, DailyInput input, DocumentConfig dailyDoc)
    {
        if (!File.Exists(prevFilePath)) return;

        using var doc   = CellHelper.OpenReadOnly(prevFilePath);
        var       wbPart = doc.WorkbookPart!;
        var       wsPart = CellHelper.FindWorksheetPart(wbPart, dailyDoc.SheetName);
        if (wsPart == null) return;

        // 누계 = 전일 칸 + 금일 칸. 전일 열이 없으면 누계 수식 칸(금일 열 + 2)의 저장된 값.
        int Cumulative(string? prevCol, string todayCol, int row)
        {
            if (!string.IsNullOrEmpty(prevCol))
                return (CellHelper.GetCellInt(wsPart, $"{prevCol}{row}") ?? 0)
                     + (CellHelper.GetCellInt(wsPart, $"{todayCol}{row}") ?? 0);
            string cumCol = IndexToColLetter(CellHelper.ColLetterToIndex(todayCol) + 2);
            return CellHelper.GetCellInt(wsPart, $"{cumCol}{row}") ?? 0;
        }

        // ── equipment_table 필드 찾기 ────────────────────────────────
        var eqField = dailyDoc.Fields.FirstOrDefault(f => f.Type == FieldTypes.EquipmentTable);
        if (eqField != null && !string.IsNullOrEmpty(eqField.NameCol) && !string.IsNullOrEmpty(eqField.TodayCol))
        {
            var used = new HashSet<EquipmentEntry>();
            for (int row = eqField.StartRow; row <= eqField.EndRow; row++)
            {
                string? cellName = CellHelper.GetCellValue(wsPart, $"{eqField.NameCol}{row}");
                if (string.IsNullOrWhiteSpace(cellName)) continue;
                string cellSpec = string.IsNullOrEmpty(eqField.SpecCol) ? ""
                                : CellHelper.GetCellValue(wsPart, $"{eqField.SpecCol}{row}") ?? "";

                // 이름이 같은 장비가 여러 줄이면(예: 덤프 15m3 / 덤프 25t) 규격으로 가르고, 한 줄에 한 항목만 짝짓는다
                int idx = FindRow(input.Equipment, e => (e.Name, e.Spec), cellName, cellSpec, i => used.Contains(input.Equipment[i]));
                if (idx < 0) continue;
                var entry = input.Equipment[idx];
                used.Add(entry);

                entry.PrevDayCount       = Cumulative(eqField.PrevDayCol, eqField.TodayCol, row);
                entry.LastFileTodayCount = CellHelper.GetCellInt(wsPart, $"{eqField.TodayCol}{row}") ?? 0;
            }
        }

        // ── personnel_table 필드 찾기 ────────────────────────────────
        var pnField = dailyDoc.Fields.FirstOrDefault(f => f.Type == FieldTypes.PersonnelTable);
        if (pnField != null && !string.IsNullOrEmpty(pnField.RoleCol) && !string.IsNullOrEmpty(pnField.TodayCol))
        {
            var people = input.GetPersonnelFor(includesSujangjae: true);   // 수장재 줄도 같은 표에 있을 수 있다
            var used   = new HashSet<PersonnelEntry>();
            for (int row = pnField.StartRow; row <= pnField.EndRow; row++)
            {
                string? cellRole = CellHelper.GetCellValue(wsPart, $"{pnField.RoleCol}{row}");
                if (string.IsNullOrWhiteSpace(cellRole)) continue;

                int idx = FindRow(people, p => (p.Role, ""), cellRole, "", i => used.Contains(people[i]));
                if (idx < 0) continue;
                var entry = people[idx];
                used.Add(entry);

                entry.PrevCumulativeCount = Cumulative(pnField.PrevCumulativeCol, pnField.TodayCol, row);
                entry.LastFileTodayCount  = CellHelper.GetCellInt(wsPart, $"{pnField.TodayCol}{row}") ?? 0;
            }
        }
    }

    /// <summary>
    /// date 보다 앞선 날짜 중 가장 최근에 만든 파일을 찾는다 (주말·쉬는 날 건너뜀).
    /// 최대 maxDays 일 전까지 본다. 없으면 null.
    /// </summary>
    public static (string Path, DateTime Date)? FindPreviousFile(
        DateTime date, string outputPattern, string baseDir, string outputFolder, int maxDays = 60)
    {
        // 폴더 목록을 한 번만 읽어 두고 비교한다 (드라이브 폴더에서 파일마다 확인하면 느리다)
        string dir = Path.GetFullPath(Path.Combine(baseDir, outputFolder));
        if (!Directory.Exists(dir)) return null;
        var names = new HashSet<string>(
            Directory.EnumerateFiles(dir).Select(f => Path.GetFileName(f)), StringComparer.OrdinalIgnoreCase);

        for (int i = 1; i <= maxDays; i++)
        {
            DateTime d    = date.Date.AddDays(-i);
            string   path = Path.GetFullPath(BuildOutputPath(d, outputPattern, baseDir, outputFolder));
            bool exists = string.Equals(Path.GetDirectoryName(path), dir, StringComparison.OrdinalIgnoreCase)
                ? names.Contains(Path.GetFileName(path))
                : File.Exists(path);   // 파일 이름에 하위 폴더가 들어간 패턴
            if (exists) return (path, d);
        }
        return null;
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
    /// <param name="photoLoader">
    /// (사진 경로, 너비 px, 높이 px) → JPEG 바이트. photo 필드가 있을 때 사진을 칸 크기에 맞춰 줄이는 함수.
    /// (사진 처리는 화면 쪽(WPF) 기능이라 App 에서 넘겨받는다)
    /// </param>
    /// <param name="notes">사용자에게 보여 줄 참고 사항(빈 칸 등)을 모을 목록 (선택)</param>
    /// <param name="problems">주면 저장 후 자동 검증(OutputVerifier)을 하고 찾은 문제를 여기에 담는다 (선택)</param>
    public static string Generate(DailyInput input, DocumentConfig docConfig,
                                  string baseDir, string outputFolder,
                                  Func<string, int, int, byte[]>? photoLoader = null,
                                  List<string>? notes = null,
                                  List<string>? problems = null)
    {
        // ── 0. 문서 날짜 (매주 문서는 설정에 따라 그 주 정한 요일 날짜) ──
        DateTime docDate = DocumentSchedule.DocumentDate(docConfig, input.Date);
        if (docDate != input.Date.Date) input = input.WithDate(docDate);

        // 명단과 연결된 표는 금일을 명단에서 센다 (화면에서도 하지만, 엔진만 쓸 때도 맞게)
        BlockEngine.ApplyRosterCounts(input, docConfig.Fields, notes);

        // ── 1. 경로 계산 ─────────────────────────────────────────────
        string templatePath = Path.Combine(baseDir, docConfig.TemplateFile);
        string outputPath   = BuildOutputPath(input.Date, docConfig.OutputPattern, baseDir, outputFolder);

        // 바탕 파일: 앞 날짜 파일이 있으면 그것 (엑셀만), 없으면 등록한 서식
        string basePath = BaseFileFor(input.Date, docConfig, baseDir, outputFolder, out var prevDate);
        if (!File.Exists(basePath))
            throw new FileNotFoundException(
                $"템플릿 파일을 찾을 수 없습니다.\n경로: {templatePath}");
        if (prevDate != null)
            notes?.Add($"{docConfig.Name}: {prevDate:MM'/'dd} 파일을 바탕으로 만들었습니다 (지정하지 않은 칸은 그날 값 그대로)");

        // 같은 날짜 파일이 이미 있으면 덮어쓰기 전에 "백업" 폴더에 남긴다
        string? backup = OutputBackup.BackupIfExists(outputPath);
        if (backup != null)
            notes?.Add($"{docConfig.Name}: 앞의 파일은 백업했습니다 → {OutputBackup.FolderName}\\{Path.GetFileName(backup)}");

        // 한글(hwpx) 서식은 별도 엔진으로 처리
        if (HwpxWriter.IsHwpx(templatePath))
        {
            EnsureDirectoryExists(outputPath);
            HwpxWriter.Generate(templatePath, outputPath, docConfig, input, notes);
            problems?.AddRange(OutputVerifier.Verify(templatePath, outputPath, docConfig));
            return outputPath;
        }

        // 식 블록의 [어제 칸] 은 앞 날짜 파일에서 미리 읽는다 (덮어쓰기 전에)
        var yesterday = ReadYesterdayForFormulas(input.Date, docConfig, baseDir, outputFolder, notes);

        // ── 2. 바탕 파일 복사 — 서식·앞 날짜 파일은 절대 수정하지 않는다 ──
        EnsureDirectoryExists(outputPath);
        File.Copy(basePath, outputPath, overwrite: true);

        // ── 3~5. 복사본을 열어 칸을 채우고 저장 (using 블록이 끝나면 파일이 닫힌다) ──
        using (var doc = SpreadsheetDocument.Open(outputPath, isEditable: true))
        {
            var wbPart = doc.WorkbookPart!;

            // 시트마다 fields 순회 (같은 모양 시트 여러 개면 모두 채운다). 시트를 따로 정한 항목은 그 시트에만.
            foreach (string sheet in docConfig.TargetSheets())
            {
                var wsPart = CellHelper.FindWorksheetPart(wbPart, sheet)
                             ?? throw new InvalidOperationException(
                                    $"시트 '{sheet}'를 찾을 수 없습니다.\n파일: {outputPath}");

                var ctx = new WriteContext(input, docConfig, sheet, doc, wsPart, photoLoader, notes, yesterday);
                foreach (var field in FormulasLast(docConfig.Fields.Where(f => string.IsNullOrWhiteSpace(f.Sheet))))
                    WriteField(field, ctx);
            }
            foreach (var group in FormulasLast(docConfig.Fields.Where(f => !string.IsNullOrWhiteSpace(f.Sheet))).GroupBy(f => f.Sheet!.Trim()))
            {
                var wsPart = CellHelper.FindWorksheetPart(wbPart, group.Key);
                if (wsPart == null)
                {
                    notes?.Add($"{docConfig.Name}: 시트 '{group.Key}' 가 없어 {string.Join(", ", group.Select(f => f.Label))} 을(를) 쓰지 못함");
                    continue;
                }
                var ctx = new WriteContext(input, docConfig, group.Key, doc, wsPart, photoLoader, notes, yesterday);
                foreach (var field in group) WriteField(field, ctx);
            }

            // 수식 재계산 플래그 설정 (누계·합계 수식 보장)
            CellHelper.SetFullCalcOnLoad(wbPart);
            wbPart.Workbook.Save();
        }

        // ── 6. 저장 후 자동 검증: 닫은 파일을 다시 열어 바탕 파일과 비교 ──
        problems?.AddRange(OutputVerifier.Verify(basePath, outputPath, docConfig));

        return outputPath;
    }

    /// <summary>
    /// 이 날짜 파일을 만들 때 복사할 바탕 파일.
    /// UsePreviousAsBase 이고 엑셀 서식이면 앞 날짜 파일(있으면), 아니면 등록한 서식.
    /// </summary>
    public static string BaseFileFor(DateTime date, DocumentConfig doc, string baseDir, string outputFolder,
                                     out DateTime? prevDate)
    {
        prevDate = null;
        string templatePath = Path.Combine(baseDir, doc.TemplateFile);
        if (!doc.UsePreviousAsBase || HwpxWriter.IsHwpx(templatePath)
            || !doc.OutputPattern.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            return templatePath;

        var prev = FindPreviousFile(date, doc.OutputPattern, baseDir, outputFolder);
        if (prev == null) return templatePath;
        prevDate = prev.Value.Date;
        return prev.Value.Path;
    }

    /// <summary>식은 다른 항목을 다 쓴 뒤에 계산한다 ([H20] 이 오늘 쓴 값을 읽도록). 나머지 순서는 그대로.</summary>
    private static IEnumerable<FieldConfig> FormulasLast(IEnumerable<FieldConfig> fields) =>
        fields.OrderBy(f => f.Type == BlockTypes.Formula ? 1 : 0);

    /// <summary>식 블록의 [어제 칸] 값. 앞 날짜 파일이 없으면 빈 값(0)으로 계산하고 알린다.</summary>
    private static Dictionary<string, string> ReadYesterdayForFormulas(DateTime date, DocumentConfig doc, string baseDir,
                                                                       string outputFolder, List<string>? notes)
    {
        bool needs = doc.Fields.Any(f => f.Type == BlockTypes.Formula
                                         && BlockFormula.Refs(f.Formula).Any(r => r.Kind == BlockFormula.KindYesterday));
        if (!needs) return [];
        var prev = FindPreviousFile(date, doc.OutputPattern, baseDir, outputFolder);
        if (prev == null)
        {
            notes?.Add($"{doc.Name}: 어제 파일이 없어 식의 [어제 칸] 을 0(빈 칸)으로 계산했습니다");
            return [];
        }
        return BlockEngine.ReadYesterdayCells(prev.Value.Path, doc);
    }

    /// <summary>필드 하나를 쓸 때 필요한 정보 묶음</summary>
    private sealed record WriteContext(
        DailyInput Input, DocumentConfig Doc, string Sheet,
        SpreadsheetDocument Package, WorksheetPart Ws,
        Func<string, int, int, byte[]>? PhotoLoader, List<string>? Notes,
        IReadOnlyDictionary<string, string>? Yesterday = null);

    /// <summary>
    /// 단순 값 필드(날짜·텍스트·직접 입력 등)의 글자 값을 구한다. 엑셀과 한글 엔진이 함께 쓴다.
    /// 값이 없는 타입이면 null.
    /// </summary>
    public static string? GetSimpleValue(FieldConfig field, DailyInput input, string docId, string sheet) =>
        field.Type switch
        {
            FieldTypes.DateKorean           => FormatKoreanDate(input.Date),
            FieldTypes.DateSimple           => input.Date.ToString("yyyy.MM.dd"),
            FieldTypes.Text                 => GetTextValue(field.Id, input),
            FieldTypes.Multiline            => GetMultilineValue(field.Id, input),
            FieldTypes.ActiveEquipmentNames => string.Join(", ",
                                                   input.Equipment.Where(e => e.TodayCount > 0).Select(e => e.Name)),
            FieldTypes.TbmPersonnel         => input.TotalPersonnelWithSujangjae.ToString(),
            FieldTypes.AttendanceRate       => "100",
            FieldTypes.Input                => WithSuffix(input.GetDocValue(docId, sheet, field.Id), field.Suffix),
            _ when BlockTypes.IsBlock(field.Type) => BlockEngine.TextFor(field, input),
            _                               => null,
        };

    /// <summary>값 뒤에 단위를 붙인다. 비어 있거나 이미 붙어 있으면 그대로.</summary>
    public static string WithSuffix(string value, string? suffix)
    {
        value = value.Trim();
        if (value.Length == 0 || string.IsNullOrEmpty(suffix)) return value;
        return value.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) ? value : value + suffix;
    }

    /// <summary>문자열을 쓰되, 수식 셀이면 건드리지 않는다.</summary>
    private static void WriteString(WriteContext c, string? cell, string value)
    {
        if (string.IsNullOrEmpty(cell) || CellHelper.HasFormula(c.Ws, cell)) return;
        CellHelper.SetCellString(c.Package, c.Ws, cell, value);
    }

    /// <summary>숫자를 쓰되, 수식 셀이면 건드리지 않는다 (예: 계획인원 칸이 다른 칸을 참조하는 서식).</summary>
    private static void WriteInt(WriteContext c, string? cell, int value)
    {
        if (string.IsNullOrEmpty(cell) || CellHelper.HasFormula(c.Ws, cell)) return;
        CellHelper.SetCellInt(c.Ws, cell, value);
    }

    // ── 타입별 처리 ───────────────────────────────────────────────────

    /// <summary>
    /// 단일 FieldConfig 를 타입에 따라 처리한다.
    /// docConfig.IncludesSujangjae 로 인원 테이블 범위를 결정한다.
    /// </summary>
    private static void WriteField(FieldConfig field, WriteContext c)
    {
        var input     = c.Input;
        var docConfig = c.Doc;
        var wsPart    = c.Ws;

        // 블록(사용자가 규칙을 고른 범용 항목)은 BlockEngine 이 쓴다
        if (BlockTypes.IsBlock(field.Type))
        {
            BlockEngine.Write(field, c.Package, wsPart, input, c.Sheet, c.Notes, c.PhotoLoader, c.Yesterday);
            return;
        }

        switch (field.Type)
        {
            // ── 날짜·텍스트: 수식 셀(다른 시트 날짜를 참조하는 칸 등)은 건드리지 않는다 ──
            case FieldTypes.DateKorean:
            case FieldTypes.DateSimple:
            case FieldTypes.Text:
                WriteString(c, field.Cell, GetSimpleValue(field, input, docConfig.Id, c.Sheet) ?? "");
                break;

            // ── 직접 입력 값 (예: 측정값 "55.3" → "55.3dB") ─────────
            case FieldTypes.Input:
            {
                string v = GetSimpleValue(field, input, docConfig.Id, c.Sheet) ?? "";
                if (v.Length > 0) WriteString(c, field.Cell, v);
                else c.Notes?.Add($"[{c.Sheet}] {field.Label}: 입력 없음 → 바탕 파일 값 그대로");
                break;
            }

            // ── 사진 칸 (+ 짝이 되는 측정값 칸) ─────────────────────
            case FieldTypes.Photo:
                WritePhoto(field, c);
                break;

            // ── 그림·글상자 위치 ─────────────────────────────────────
            case FieldTypes.ShapePosition:
            {
                string name   = input.GetDocValue(docConfig.Id, c.Sheet, field.Id);
                var    preset = field.Presets?.FirstOrDefault(p => p.Name == name);
                if (preset != null)
                {
                    if (ExcelDrawingHelper.ApplyPreset(wsPart, preset) < preset.Shapes.Count)
                        c.Notes?.Add($"[{c.Sheet}] {field.Label}: 일부 그림을 이 시트에서 찾지 못함");
                }
                else if (name.Length > 0) c.Notes?.Add($"[{c.Sheet}] {field.Label}: 저장된 위치 '{name}' 없음");
                break;
            }

            // ── 멀티라인 텍스트 → 셀 범위에 나눠 쓰기 ───────────────
            case FieldTypes.Multiline:
                if (field.Cells is { Count: > 0 })
                    WriteMultiline(c, GetMultilineValue(field.Id, input), field.Cells);
                break;

            // ── 장비 투입현황 테이블 ──────────────────────────────────
            case FieldTypes.EquipmentTable:
                WriteEquipmentTable(c, input.Equipment, field);
                break;

            // ── 인원 투입현황 테이블 ─────────────────────────────────
            // docConfig.IncludesSujangjae = true 이면 수장재 포함 목록 사용
            case FieldTypes.PersonnelTable:
                WritePersonnelTable(c, input.GetPersonnelFor(docConfig.IncludesSujangjae), field);
                break;

            // ── 금일 대수 > 0 장비명 쉼표 연결 ──────────────────────
            case FieldTypes.ActiveEquipmentNames:
                WriteString(c, field.Cell, GetSimpleValue(field, input, docConfig.Id, c.Sheet) ?? "");
                break;

            // ── TBM 인원 — 수장재 포함 합계 사용 ────────────────────
            case FieldTypes.TbmPersonnel:
                WriteInt(c, field.Cell, input.TotalPersonnelWithSujangjae);
                break;

            // ── 참석율 — 계획 = 실시이므로 항상 100 ─────────────────
            case FieldTypes.AttendanceRate:
                WriteInt(c, field.Cell, 100);
                break;
        }
    }

    /// <summary>
    /// 사진 칸: 그 칸에 놓인 기존 그림을 새 사진으로 바꾸고, 짝 값 칸(ValueCell)이 있으면 측정값도 쓴다.
    /// 사진이 없으면 서식의 그림을 그대로 둔다.
    /// </summary>
    private static void WritePhoto(FieldConfig field, WriteContext c)
    {
        string docId = c.Doc.Id;
        string path  = c.Input.GetDocValue(docId, c.Sheet, field.Id);

        if (!string.IsNullOrEmpty(field.Cell))
        {
            if (path.Length > 0 && File.Exists(path) && c.PhotoLoader != null)
            {
                bool ok = ExcelDrawingHelper.ReplacePictureAt(c.Ws, field.Cell,
                    (w, h) => c.PhotoLoader(path, w, h));
                if (!ok) c.Notes?.Add($"[{c.Sheet}] {field.Label}: {field.Cell} 칸에 그림이 없어 사진을 넣지 못함");
            }
            else
                c.Notes?.Add($"[{c.Sheet}] {field.Label}: 사진 없음 → 바탕 파일(서식·앞 날짜) 사진 그대로");
        }

        if (!string.IsNullOrEmpty(field.ValueCell))
        {
            string v = WithSuffix(c.Input.GetDocValue(docId, c.Sheet, field.Id + DailyInput.ValueSuffixKey), field.Suffix);
            if (v.Length > 0) WriteString(c, field.ValueCell, v);
            else c.Notes?.Add($"[{c.Sheet}] {field.Label}: 측정값 없음 → 바탕 파일 값 그대로");
        }
    }

    /// <summary>
    /// 멀티라인 텍스트를 셀 범위에 나눠 쓴다.
    /// 텍스트 줄 수가 셀 범위보다 적으면 나머지 셀은 빈 문자열로 클리어한다.
    /// </summary>
    private static void WriteMultiline(WriteContext c, string text, List<string> cells)
    {
        string[] lines = text.Split(["\r\n", "\n"], StringSplitOptions.None);

        for (int i = 0; i < cells.Count; i++)
            WriteString(c, cells[i], i < lines.Length ? lines[i] : "");   // 수식 칸은 건너뛴다
    }

    /// <summary>장비 투입현황: 장비명 열로 줄을 찾아 전일·금일을 쓴다.</summary>
    private static void WriteEquipmentTable(WriteContext c, List<EquipmentEntry> equipment, FieldConfig field) =>
        WriteCountTable(c, field, field.NameCol, field.SpecCol, field.PrevDayCol,
            equipment.Select(e => new CountRow(e.Name, e.Spec, e.PrevDayCount, e.TodayCount)).ToList());

    /// <summary>인원 투입현황: 직종 열로 줄을 찾아 전일누계·금일을 쓴다.</summary>
    private static void WritePersonnelTable(WriteContext c, List<PersonnelEntry> personnel, FieldConfig field) =>
        WriteCountTable(c, field, field.RoleCol, null, field.PrevCumulativeCol,
            personnel.Select(p => new CountRow(p.Role, "", p.PrevCumulativeCount, p.TodayCount)).ToList());

    private sealed record CountRow(string Key, string Spec, int Prev, int Today);

    /// <summary>
    /// 장비·인원 표 공통.
    /// - 화면에 있는 줄: 전일·금일을 쓴다.
    /// - 파일에만 있는 줄(화면에서 지운 줄): 금일을 0 으로, 전일은 "전일 + 어제 금일" 로 넘겨 누계가 끊기지 않게 한다.
    /// - 화면에만 있는 줄(새로 추가한 장비·직종): 표의 빈 줄에 이름(·규격)과 값을 쓴다.
    /// - 누계 열(금일 + 2)은 수식이면 두고, 숫자로 적혀 있으면 전일 + 금일 을 써 준다.
    /// </summary>
    private static void WriteCountTable(WriteContext c, FieldConfig field, string? keyCol, string? specCol,
                                        string? prevCol, List<CountRow> rows)
    {
        string? todayCol = field.TodayCol;
        if (string.IsNullOrEmpty(keyCol) || string.IsNullOrEmpty(prevCol) || string.IsNullOrEmpty(todayCol))
            return;
        var ws = c.Ws;
        string cumCol = IndexToColLetter(CellHelper.ColLetterToIndex(todayCol) + 2);

        void SetCounts(int row, int prev, int today)
        {
            if (!CellHelper.HasFormula(ws, $"{prevCol}{row}"))  CellHelper.SetCellInt(ws, $"{prevCol}{row}", prev);
            if (!CellHelper.HasFormula(ws, $"{todayCol}{row}")) CellHelper.SetCellInt(ws, $"{todayCol}{row}", today);
            if (!CellHelper.HasFormula(ws, $"{cumCol}{row}"))   CellHelper.SetCellInt(ws, $"{cumCol}{row}", prev + today);
        }

        var used      = new bool[rows.Count];
        var emptyRows = new List<int>();
        for (int row = field.StartRow; row <= field.EndRow; row++)
        {
            string? name = CellHelper.GetCellValue(ws, $"{keyCol}{row}");
            if (string.IsNullOrWhiteSpace(name)) { emptyRows.Add(row); continue; }

            string spec = string.IsNullOrEmpty(specCol) ? "" : CellHelper.GetCellValue(ws, $"{specCol}{row}") ?? "";
            int idx = FindRow(rows, r => (r.Key, r.Spec), name, spec, i => used[i]);

            if (idx >= 0)
            {
                used[idx] = true;
                SetCounts(row, rows[idx].Prev, rows[idx].Today);
            }
            else
            {
                int p = CellHelper.GetCellInt(ws, $"{prevCol}{row}") ?? 0;
                int t = CellHelper.GetCellInt(ws, $"{todayCol}{row}") ?? 0;
                SetCounts(row, p + t, 0);
            }
        }

        for (int i = 0; i < rows.Count; i++)
        {
            if (used[i] || string.IsNullOrWhiteSpace(rows[i].Key)) continue;
            if (emptyRows.Count == 0)
            {
                c.Notes?.Add($"[{c.Sheet}] {field.Label}: 빈 줄이 없어 '{rows[i].Key}' 을(를) 넣지 못함");
                continue;
            }
            int row = emptyRows[0];
            emptyRows.RemoveAt(0);
            CellHelper.SetCellString(c.Package, ws, $"{keyCol}{row}", rows[i].Key);
            if (!string.IsNullOrEmpty(specCol) && rows[i].Spec.Length > 0)
                CellHelper.SetCellString(c.Package, ws, $"{specCol}{row}", rows[i].Spec);
            SetCounts(row, rows[i].Prev, rows[i].Today);
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
    /// <summary>이 문서를 date 에 만들면 저장될 경로 (매주 문서의 날짜 규칙 반영). 화면의 덮어쓰기 확인 등에 쓴다.</summary>
    public static string OutputPathFor(DocumentConfig doc, DateTime date, string baseDir, string outputFolder) =>
        BuildOutputPath(DocumentSchedule.DocumentDate(doc, date), doc.OutputPattern, baseDir, outputFolder);

    public static string BuildOutputPath(DateTime date, string outputPattern,
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
    public static string FormatKoreanDate(DateTime date)
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
    /// 표의 한 줄(이름·규격)에 짝지을 화면 항목을 찾는다. 이미 짝지은 항목은 건너뛴다.
    /// 이름과 규격이 모두 같은 항목을 먼저, 없으면 이름만 같은 항목 중 첫 번째.
    /// (같은 이름 장비가 여러 줄이어도 화면 순서와 상관없이 제 줄을 찾게 한다)
    /// </summary>
    private static int FindRow<T>(IReadOnlyList<T> items, Func<T, (string Name, string Spec)> key,
                                  string name, string spec, Func<int, bool> isUsed)
    {
        int nameOnly = -1;
        for (int i = 0; i < items.Count; i++)
        {
            if (isUsed(i)) continue;
            var (n, sp) = key(items[i]);
            if (NormName(n) != NormName(name)) continue;
            if (NormName(sp) == NormName(spec)) return i;
            if (nameOnly < 0) nameOnly = i;
        }
        return nameOnly;
    }

    /// <summary>
    /// 셀 텍스트 비교용 정규화: 앞뒤 공백 제거, 대소문자 소문자 통일.
    /// "B/H 380 " vs "B/H 380" 처럼 공백 차이를 무시하기 위해 사용.
    /// </summary>
    private static string NormName(string s)
        => s.Trim().ToLowerInvariant();
}
