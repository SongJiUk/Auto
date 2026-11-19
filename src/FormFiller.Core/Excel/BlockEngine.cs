using DocumentFormat.OpenXml.Packaging;
using FormFiller.Core.Config;
using FormFiller.Core.Models;

namespace FormFiller.Core.Excel;

/// <summary>
/// 블록(BlockTypes) 쓰기·읽기·계산.
/// 블록은 "칸 위치 + 사용자가 고른 규칙"만 가지고 있고, 뜻(날씨·장비 등)은 모른다.
///
/// [흐름]
/// 1. 화면이 ReadValues 로 앞 날짜 파일(없으면 서식)에서 처음 값을 읽어 보여 준다 (전일 누계, 어제 작업 등)
/// 2. 사용자가 고친 값이 DailyInput.Blocks 에 담겨 온다
/// 3. Generate 가 블록마다 Write 를 부른다
/// </summary>
public static class BlockEngine
{
    // ════════════════════════════════════════════════════════════════
    // 쓰기
    // ════════════════════════════════════════════════════════════════

    /// <param name="photoLoader">(사진 경로, 너비 px, 높이 px) → JPEG 바이트. 사진 구역 블록에서만 쓴다.</param>
    public static void Write(FieldConfig f, SpreadsheetDocument pkg, WorksheetPart ws, DailyInput input,
                             string sheet, List<string>? notes, Func<string, int, int, byte[]>? photoLoader = null)
    {
        string key = BlockTypes.KeyOf(f);
        var v = input.Block(key);
        switch (f.Type)
        {
            case BlockTypes.Value:
            case BlockTypes.Date:
            case BlockTypes.Calc:
            {
                string? text = TextFor(f, input);
                if (text == null) return;                                  // 어제 값 그대로
                if (text.Length == 0 && f.Type == BlockTypes.Value)
                {
                    notes?.Add($"[{sheet}] {f.Label}: 입력 없음 → 바탕 파일 값 그대로");
                    return;
                }
                bool asNumber = f.Type == BlockTypes.Calc ? f.CalcKind != BlockTypes.CalcNames
                              : f.Type == BlockTypes.Value && string.IsNullOrEmpty(f.Suffix);
                WriteText(pkg, ws, f.Cell, text, asNumber);
                break;
            }

            case BlockTypes.Lines:
                if (f.Cells is not { Count: > 0 }) return;
                var lines = PrefixLines(v.Lines, f.Prefix);
                for (int i = 0; i < f.Cells.Count; i++)
                    WriteText(pkg, ws, f.Cells[i], i < lines.Count ? lines[i] : "", asNumber: false);
                if (lines.Count > f.Cells.Count)
                    notes?.Add($"[{sheet}] {f.Label}: {lines.Count}줄 중 칸이 {f.Cells.Count}개라 뒤 {lines.Count - f.Cells.Count}줄은 넣지 못함");
                break;

            case BlockTypes.Counter:
                WriteNumber(ws, f.Cell, v.Today);
                WriteNumber(ws, f.PrevCell, v.Prev);
                WriteNumber(ws, f.TotalCell, v.Prev + v.Today);
                break;

            case BlockTypes.Table:
                WriteTable(f, pkg, ws, v.Rows, sheet, notes);
                break;

            case BlockTypes.Roster:
                WriteRoster(f, pkg, ws, v.People, sheet, notes);
                break;

            case BlockTypes.Photos:
                PhotoAreaWriter.Write(f, pkg, ws, v, sheet, photoLoader, notes);
                break;
        }
    }

    /// <summary>
    /// 칸 하나짜리 블록의 글자 값 (엑셀·한글 공통). null 이면 "쓰지 않음"(어제 값 그대로).
    /// </summary>
    public static string? TextFor(FieldConfig f, DailyInput input)
    {
        var v = input.Block(BlockTypes.KeyOf(f));
        return f.Type switch
        {
            BlockTypes.Value => (f.ValueSource ?? BlockTypes.SourceInput) switch
            {
                BlockTypes.SourceKeep  => null,
                BlockTypes.SourceFixed => f.FixedValue ?? "",
                _                      => GenericExcelWriter.WithSuffix(v.Text, f.Suffix),
            },
            BlockTypes.Date    => BlockTypes.FormatDate(f.DateFormat, input.Date),
            BlockTypes.Lines   => string.Join("\n", PrefixLines(v.Lines, f.Prefix)),
            BlockTypes.Counter => Num(v.Today),
            BlockTypes.Calc    => Calc(f, input),
            _                  => null,
        };
    }

    /// <summary>줄마다 앞 글자를 붙인다. 빈 줄·앞 글자만 있는 줄은 뺀다.</summary>
    public static List<string> PrefixLines(IEnumerable<string> lines, string? prefix)
    {
        var result = new List<string>();
        string p = prefix ?? "";
        string bare = p.Trim();
        foreach (var raw in lines)
        {
            string line = raw.TrimEnd();
            if (line.Trim().Length == 0 || (bare.Length > 0 && line.Trim() == bare)) continue;
            result.Add(p.Length == 0 || line.StartsWith(p) || (bare.Length > 0 && line.TrimStart().StartsWith(bare))
                ? line : p + line);
        }
        return result;
    }

    private static string Num(decimal d) => d.ToString("0.############", System.Globalization.CultureInfo.InvariantCulture);

    /// <summary>수식 칸은 건드리지 않고 글자(또는 숫자)를 쓴다.</summary>
    private static void WriteText(SpreadsheetDocument pkg, WorksheetPart ws, string? cell, string text, bool asNumber)
    {
        if (string.IsNullOrEmpty(cell) || CellHelper.HasFormula(ws, cell)) return;
        if (asNumber && BlockTypes.ParseNumber(text) is { } n) CellHelper.SetCellNumber(ws, cell, n);
        else CellHelper.SetCellString(pkg, ws, cell, text);
    }

    private static void WriteNumber(WorksheetPart ws, string? cell, decimal value)
    {
        if (string.IsNullOrEmpty(cell) || CellHelper.HasFormula(ws, cell)) return;
        CellHelper.SetCellNumber(ws, cell, value);
    }

    // ── 표 ───────────────────────────────────────────────────────────

    /// <summary>
    /// 표: 이름(+규격)으로 줄을 찾아 전일·금일·누계를 쓴다.
    /// - 화면에 없는 파일 줄: 금일 0, 전일은 "전일 + 어제 금일" 로 넘겨 누계가 끊기지 않게
    /// - 화면에만 있는 줄: 빈 줄에 이름·규격과 값
    /// - 누계 칸은 수식이면 두고, 숫자면 전일 + 금일
    /// </summary>
    private static void WriteTable(FieldConfig f, SpreadsheetDocument pkg, WorksheetPart ws, List<TableRowValue> rows,
                                   string sheet, List<string>? notes)
    {
        string? nameCol  = BlockTypes.ColOf(f, BlockTypes.RoleName);
        string? specCol  = BlockTypes.ColOf(f, BlockTypes.RoleSpec);
        string? prevCol  = BlockTypes.ColOf(f, BlockTypes.RolePrev);
        string? todayCol = BlockTypes.ColOf(f, BlockTypes.RoleToday);
        string? totalCol = BlockTypes.ColOf(f, BlockTypes.RoleTotal);
        if (nameCol == null || todayCol == null || f.StartRow <= 0) return;

        void Set(int row, decimal prev, decimal today)
        {
            WriteNumber(ws, prevCol  is null ? null : $"{prevCol}{row}", prev);
            WriteNumber(ws, $"{todayCol}{row}", today);
            WriteNumber(ws, totalCol is null ? null : $"{totalCol}{row}", prev + today);
        }

        var used = new bool[rows.Count];
        var empty = new List<int>();
        for (int row = f.StartRow; row <= f.EndRow; row++)
        {
            string? name = CellHelper.GetCellValue(ws, $"{nameCol}{row}");
            if (string.IsNullOrWhiteSpace(name)) { empty.Add(row); continue; }
            string spec = specCol is null ? "" : CellHelper.GetCellValue(ws, $"{specCol}{row}") ?? "";

            int idx = FindRow(rows, name, spec, used);
            if (idx >= 0) { used[idx] = true; Set(row, rows[idx].Prev, rows[idx].Today); }
            else
            {
                // 화면에서 지운 줄: 누계가 끊기지 않게 어제까지 값을 전일로
                decimal cum = CumulativeAt(ws, prevCol, todayCol, totalCol, row);
                Set(row, cum, 0);
            }
        }

        for (int i = 0; i < rows.Count; i++)
        {
            if (used[i] || string.IsNullOrWhiteSpace(rows[i].Name)) continue;
            if (empty.Count == 0) { notes?.Add($"[{sheet}] {f.Label}: 빈 줄이 없어 '{rows[i].Name}' 을(를) 넣지 못함"); continue; }
            int row = empty[0]; empty.RemoveAt(0);
            WriteText(pkg, ws, $"{nameCol}{row}", rows[i].Name, asNumber: false);
            if (specCol != null && rows[i].Spec.Length > 0) WriteText(pkg, ws, $"{specCol}{row}", rows[i].Spec, asNumber: false);
            Set(row, rows[i].Prev, rows[i].Today);
        }
    }

    /// <summary>
    /// 그 줄의 "지금까지 누계": 전일 열이 있으면 전일 + 금일 (프로그램이 만든 파일은 누계 수식의 저장값이 낡았을 수 있어서),
    /// 없으면 누계 열 값, 그것도 없으면 금일.
    /// </summary>
    private static decimal CumulativeAt(WorksheetPart ws, string? prevCol, string todayCol, string? totalCol, int row)
    {
        decimal today = CellHelper.GetCellDecimal(ws, $"{todayCol}{row}") ?? 0;
        if (prevCol != null) return (CellHelper.GetCellDecimal(ws, $"{prevCol}{row}") ?? 0) + today;
        if (totalCol != null) return CellHelper.GetCellDecimal(ws, $"{totalCol}{row}") ?? today;
        return today;
    }

    private static string Norm(string s) => s.Trim().ToLowerInvariant();

    /// <summary>이름·규격이 모두 같은 줄을 먼저, 없으면 이름만 같은 줄 (이미 짝지은 줄은 건너뜀)</summary>
    private static int FindRow(List<TableRowValue> rows, string name, string spec, bool[] used)
    {
        int nameOnly = -1;
        for (int i = 0; i < rows.Count; i++)
        {
            if (used[i] || Norm(rows[i].Name) != Norm(name)) continue;
            if (Norm(rows[i].Spec) == Norm(spec)) return i;
            if (nameOnly < 0) nameOnly = i;
        }
        return nameOnly;
    }

    // ── 명단 ─────────────────────────────────────────────────────────

    /// <summary>명단: 사람을 위에서부터 한 줄씩 쓰고, 남는 줄의 직종·성명·작업내용은 비운다 (번호 열은 그대로).</summary>
    private static void WriteRoster(FieldConfig f, SpreadsheetDocument pkg, WorksheetPart ws, List<RosterEntry> people,
                                    string sheet, List<string>? notes)
    {
        string? jobCol    = BlockTypes.ColOf(f, BlockTypes.RoleJob);
        string? personCol = BlockTypes.ColOf(f, BlockTypes.RolePerson);
        string? taskCol   = BlockTypes.ColOf(f, BlockTypes.RoleTask);
        if (personCol == null || f.StartRow <= 0) return;

        int capacity = f.EndRow - f.StartRow + 1;
        for (int i = 0; i < capacity; i++)
        {
            int row = f.StartRow + i;
            var p = i < people.Count ? people[i] : null;
            void Put(string? col, string? value)
            {
                if (col == null) return;
                string cell = $"{col}{row}";
                // 남는 줄은 원래 비어 있으면 건드리지 않는다 (빈 칸을 새로 만들지 않게)
                if (value == null && string.IsNullOrEmpty(CellHelper.GetCellValue(ws, cell))) return;
                WriteText(pkg, ws, cell, value ?? "", asNumber: false);
            }
            Put(jobCol,    p?.Job);
            Put(personCol, p?.Person);
            Put(taskCol,   p?.Task);
        }
        if (people.Count > capacity)
            notes?.Add($"[{sheet}] {f.Label}: {people.Count}명 중 줄이 {capacity}개라 {people.Count - capacity}명은 넣지 못함");
    }

    // ════════════════════════════════════════════════════════════════
    // 계산
    // ════════════════════════════════════════════════════════════════

    /// <summary>계산 블록의 결과 글자</summary>
    public static string Calc(FieldConfig f, DailyInput input)
    {
        var sources = (f.Sources ?? []).Select(s => input.Block(s.Trim())).ToList();
        return (f.CalcKind ?? BlockTypes.CalcSum) switch
        {
            BlockTypes.CalcNames => string.Join(", ", sources.SelectMany(v => v.Rows).Where(r => r.Today > 0).Select(r => r.Name)),
            BlockTypes.CalcCount => Num(sources.Sum(v => v.People.Count)),
            _                    => Num(sources.Sum(NumberOf)),
        };
    }

    /// <summary>블록 값을 숫자 하나로: 표 = 금일 합계, 명단 = 인원 수, 누계 = 오늘, 값 = 글자를 숫자로</summary>
    public static decimal NumberOf(BlockValue v) =>
        v.Rows.Count > 0   ? v.Rows.Sum(r => r.Today)
      : v.People.Count > 0 ? v.People.Count
      : v.Today != 0       ? v.Today
      : BlockTypes.ParseNumber(v.Text) ?? 0;

    /// <summary>
    /// 명단과 연결된 표(RosterKey)의 금일을 명단에서 직종별로 세어 채운다.
    /// 짝짓기(RoleMap)가 없으면 줄 이름과 같은 직종을 센다. 표에 없는 직종은 notes 로 알린다.
    /// </summary>
    public static void ApplyRosterCounts(DailyInput input, IEnumerable<FieldConfig> tableFields, List<string>? notes = null)
    {
        foreach (var f in tableFields.Where(f => f.Type == BlockTypes.Table && !string.IsNullOrWhiteSpace(f.RosterKey)))
        {
            var people = input.Block(f.RosterKey!.Trim()).People;
            var table  = input.Block(BlockTypes.KeyOf(f));
            var counted = new HashSet<int>();
            foreach (var row in table.Rows)
            {
                var jobs = f.RoleMap != null && f.RoleMap.TryGetValue(row.Name, out var mapped) && mapped.Count > 0
                    ? mapped : [row.Name];
                var set = jobs.Select(Norm).ToHashSet();
                int n = 0;
                for (int i = 0; i < people.Count; i++)
                    if (set.Contains(Norm(people[i].Job))) { n++; counted.Add(i); }
                row.Today = n;
            }
            var missing = people.Where((p, i) => !counted.Contains(i) && p.Person.Length > 0)
                                .Select(p => string.IsNullOrWhiteSpace(p.Job) ? "(직종 없음)" : p.Job).Distinct().ToList();
            if (missing.Count > 0)
                notes?.Add($"{f.Label}: 명단의 직종 {string.Join(", ", missing)} 을(를) 셀 줄이 없습니다 (설정의 짝짓기 확인)");
        }
    }

    // ════════════════════════════════════════════════════════════════
    // 읽기 (화면의 처음 값)
    // ════════════════════════════════════════════════════════════════

    /// <summary>
    /// 파일(앞 날짜 파일 또는 서식)의 한 시트에서 블록 값을 읽는다.
    /// - value: 그 칸 글자 / lines: 줄 목록 / counter: Prev = 그 파일 기준 누계(→ 오늘의 전일)
    /// - table: 줄 이름·규격 + Prev = 그 파일 기준 누계, LastToday = 그 파일 금일
    /// - roster: 적힌 사람들
    /// </summary>
    public static Dictionary<string, BlockValue> ReadValues(string filePath, DocumentConfig doc, string? sheet = null)
    {
        var result = new Dictionary<string, BlockValue>();
        if (!File.Exists(filePath) || HwpxWriter.IsHwpx(filePath)) return result;

        using var pkg = CellHelper.OpenReadOnly(filePath);
        string defaultSheet = sheet ?? doc.TargetSheets().FirstOrDefault() ?? "";

        foreach (var f in doc.Fields.Where(f => BlockTypes.IsBlock(f.Type)))
        {
            // 시트를 따로 정한 항목은 그 시트에서 읽는다
            var ws = CellHelper.FindWorksheetPart(pkg.WorkbookPart!,
                         string.IsNullOrWhiteSpace(f.Sheet) ? defaultSheet : f.Sheet.Trim());
            if (ws == null) continue;
            var v = new BlockValue();
            switch (f.Type)
            {
                case BlockTypes.Value:
                    v.Text = string.IsNullOrEmpty(f.Cell) ? "" : StripSuffix(CellHelper.GetCellValue(ws, f.Cell) ?? "", f.Suffix);
                    break;
                case BlockTypes.Lines:
                    v.Lines = (f.Cells ?? []).Select(c => CellHelper.GetCellValue(ws, c) ?? "")
                                             .Where(s => s.Trim().Length > 0).ToList();
                    break;
                case BlockTypes.Counter:
                {
                    decimal today = Read(ws, f.Cell);
                    v.Prev = !string.IsNullOrEmpty(f.TotalCell) && !CellHelper.HasFormula(ws, f.TotalCell)
                        ? Read(ws, f.TotalCell)
                        : Read(ws, f.PrevCell) + today;
                    break;
                }
                case BlockTypes.Table:
                    v.Rows = ReadTableRows(f, ws);
                    break;
                case BlockTypes.Roster:
                    v.People = ReadRoster(f, ws);
                    break;
                case BlockTypes.Photos:
                    v.Photos = PhotoAreaWriter.ReadCaptions(f, ws);
                    break;
                default:
                    continue;
            }
            result.TryAdd(BlockTypes.KeyOf(f), v);
        }
        return result;
    }

    /// <summary>
    /// 앞 날짜 파일(없으면 서식)에서 읽은 값으로 "오늘 화면의 처음 값"을 만든다. 화면과 테스트가 같이 쓴다.
    /// - value(매일 입력): 어제 글자를 그대로 보여 준다 (고쳐 쓰면 됨)
    /// - lines: StartFrom 규칙 (다른 블록 → 어제 그 블록, 비었으면 어제 이 블록 / "same" → 어제 이 블록 / 없음 → 빈 칸)
    /// - counter: 전일 = 어제 누계, 오늘 = 0
    /// - table: 줄 목록 + 전일 = 어제 누계, 금일 = 0 (LastToday 는 [전일과 같게] 용)
    /// - roster: 어제 나온 사람들
    /// </summary>
    public static Dictionary<string, BlockValue> InitialValues(DocumentConfig doc, IReadOnlyDictionary<string, BlockValue> prev)
    {
        var result = new Dictionary<string, BlockValue>();
        foreach (var f in doc.Fields.Where(f => BlockTypes.IsBlock(f.Type)))
        {
            string key = BlockTypes.KeyOf(f);
            if (result.ContainsKey(key)) continue;
            prev.TryGetValue(key, out var p);
            var v = new BlockValue();
            switch (f.Type)
            {
                case BlockTypes.Value:
                    v.Text = p?.Text ?? "";
                    break;
                case BlockTypes.Lines:
                {
                    string from = f.StartFrom?.Trim() ?? "";
                    if (from.Length > 0 && from != "same" && prev.TryGetValue(from, out var other) && other.Lines.Count > 0)
                        v.Lines = [.. other.Lines];
                    else if (from.Length > 0)
                        v.Lines = [.. p?.Lines ?? []];
                    break;
                }
                case BlockTypes.Counter:
                    v.Prev = p?.Prev ?? 0;
                    break;
                case BlockTypes.Table:
                    v.Rows = (p?.Rows ?? []).Select(r => new TableRowValue
                        { Name = r.Name, Spec = r.Spec, Prev = r.Prev, LastToday = r.LastToday }).ToList();
                    break;
                case BlockTypes.Roster:
                    v.People = (p?.People ?? []).Select(x => new RosterEntry { Job = x.Job, Person = x.Person, Task = x.Task }).ToList();
                    break;
                case BlockTypes.Photos:
                    // 설명은 어제 것을 그대로, 사진은 매일 새로 고른다
                    v.Photos = (p?.Photos ?? []).Select(x => new PhotoValue { Caption = x.Caption }).ToList();
                    break;
                default:
                    continue;
            }
            result[key] = v;
        }
        return result;
    }

    private static decimal Read(WorksheetPart ws, string? cell) =>
        string.IsNullOrEmpty(cell) ? 0 : CellHelper.GetCellDecimal(ws, cell) ?? 0;

    private static string StripSuffix(string s, string? suffix) =>
        !string.IsNullOrEmpty(suffix) && s.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) ? s[..^suffix.Length] : s;

    private static List<TableRowValue> ReadTableRows(FieldConfig f, WorksheetPart ws)
    {
        var rows = new List<TableRowValue>();
        string? nameCol = BlockTypes.ColOf(f, BlockTypes.RoleName), specCol = BlockTypes.ColOf(f, BlockTypes.RoleSpec);
        string? prevCol = BlockTypes.ColOf(f, BlockTypes.RolePrev), todayCol = BlockTypes.ColOf(f, BlockTypes.RoleToday);
        string? totalCol = BlockTypes.ColOf(f, BlockTypes.RoleTotal);
        if (nameCol == null || todayCol == null) return rows;

        for (int row = f.StartRow; row <= f.EndRow; row++)
        {
            string? name = CellHelper.GetCellValue(ws, $"{nameCol}{row}")?.Trim();
            if (string.IsNullOrEmpty(name)) continue;
            // 누계 칸이 수식이면 저장값이 낡았을 수 있으니 전일 + 금일로 센다
            string? total = totalCol != null && !CellHelper.HasFormula(ws, $"{totalCol}{row}") ? totalCol : null;
            rows.Add(new TableRowValue
            {
                Name      = name,
                Spec      = specCol == null ? "" : CellHelper.GetCellValue(ws, $"{specCol}{row}")?.Trim() ?? "",
                Prev      = CumulativeAt(ws, prevCol, todayCol, total, row),
                LastToday = CellHelper.GetCellDecimal(ws, $"{todayCol}{row}") ?? 0,
            });
        }
        return rows;
    }

    private static List<RosterEntry> ReadRoster(FieldConfig f, WorksheetPart ws)
    {
        var list = new List<RosterEntry>();
        string? jobCol = BlockTypes.ColOf(f, BlockTypes.RoleJob), personCol = BlockTypes.ColOf(f, BlockTypes.RolePerson);
        string? taskCol = BlockTypes.ColOf(f, BlockTypes.RoleTask);
        if (personCol == null) return list;
        for (int row = f.StartRow; row <= f.EndRow; row++)
        {
            string person = CellHelper.GetCellValue(ws, $"{personCol}{row}")?.Trim() ?? "";
            if (person.Length == 0) continue;
            list.Add(new RosterEntry
            {
                Person = person,
                Job    = jobCol  == null ? "" : CellHelper.GetCellValue(ws, $"{jobCol}{row}")?.Trim() ?? "",
                Task   = taskCol == null ? "" : CellHelper.GetCellValue(ws, $"{taskCol}{row}")?.Trim() ?? "",
            });
        }
        return list;
    }
}
