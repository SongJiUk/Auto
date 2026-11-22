using System.Globalization;
using System.Text.Json;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Core.Models;

namespace FormFiller.Core.Phone;

// ════════════════════════════════════════════════════════════════
// 폰 화면 양식 (phone_form.json)
// PC 가 드라이브 inbox 폴더에 써 두면 폰 웹앱이 읽어서 입력 화면을 만든다.
// 폰은 엑셀을 못 읽으므로 "어제 장비 목록·사람 목록" 같은 처음 값을 PC 가 미리 풀어 둔다.
// 이름이 들어 있지만 사용자 드라이브 안에만 있다. API 키 같은 설정은 넣지 않는다.
// ════════════════════════════════════════════════════════════════

public sealed class PhoneForm
{
    public int SchemaVersion { get; init; } = PhoneDaily.SCHEMA_VERSION;
    /// <summary>처음 값을 읽은 날짜 (yyyy-MM-dd)</summary>
    public string Date { get; init; } = "";
    public string UpdatedAt { get; init; } = "";
    public string UpdatedBy { get; init; } = "";
    public List<PhoneFormDoc> Docs { get; init; } = [];
    public List<PhoneFormBlock> Blocks { get; init; } = [];
    public List<PhoneFormPerson> People { get; init; } = [];
    public List<PhoneFormDocField> DocFields { get; init; } = [];
    public List<PhoneFormSlot> Slots { get; init; } = [];
}

/// <param name="Due">오늘 만들 차례인지 (폰에서 처음부터 체크)</param>
public sealed record PhoneFormDoc(string Id, string Name, bool Due);

public sealed class PhoneFormBlock
{
    public string Key   { get; init; } = "";
    public string Label { get; init; } = "";
    /// <summary>value · lines · counter · table · roster · photos</summary>
    public string Type  { get; init; } = "";
    public List<string> Docs { get; init; } = [];
    public string Text  { get; init; } = "";
    public List<string> Lines { get; init; } = [];
    public List<PhoneFormRow> Rows { get; init; } = [];
    /// <summary>표 금일을 명단에서 세면 true (폰에서 숫자를 받지 않는다)</summary>
    public bool TodayFromRoster { get; init; }
    public int PhotoCount { get; init; }
}

/// <param name="LastToday">어제 금일 (폰에서 [어제와 같게] 용)</param>
public sealed record PhoneFormRow(string Name, string Spec, decimal LastToday);

public sealed record PhoneFormPerson(string Name, string Job, string Task, bool Yesterday);

/// <param name="Kind">photo · input · shape</param>
/// <param name="HasValue">사진과 함께 측정값을 받는지</param>
public sealed record PhoneFormDocField(string Doc, string DocName, string Sheet, string Field, string Label,
                                       string Kind, bool HasValue, string Unit, List<string> Options);

public sealed record PhoneFormSlot(string Doc, string DocName, string Sheet, int Index, string Label);

/// <summary>폰 입력 공통: 블록 처음 값 읽기, 양식 만들기</summary>
public static class PhoneDaily
{
    public const int SCHEMA_VERSION = 2;
    public const string FORM_FILE_NAME = "phone_form.json";
    private const string DATE_FORMAT = "yyyy-MM-dd";

    public static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy        = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented               = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static string DateText(DateTime d) => d.ToString(DATE_FORMAT, CultureInfo.InvariantCulture);

    /// <summary>
    /// 블록의 처음 값 (어제 누계·어제 장비 목록 등). 문서마다 앞 날짜 파일(없으면 서식)에서 읽는다.
    /// 오늘 작성 화면(MainViewModel.ReloadBlockValues)과 같은 순서·규칙.
    /// </summary>
    public static Dictionary<string, BlockValue> InitialBlocks(AppConfig cfg, string baseDir, DateTime date,
                                                              List<string>? notes = null)
    {
        var all = new Dictionary<string, BlockValue>();
        foreach (var doc in cfg.Documents.Where(d => d.Fields.Any(f => BlockTypes.IsBlock(f.Type))))
        {
            try
            {
                var prev = GenericExcelWriter.FindPreviousFile(DocumentSchedule.DocumentDate(doc, date),
                                                               doc.OutputPattern, baseDir, cfg.OutputFolder);
                string file = prev?.Path ?? Path.Combine(baseDir, doc.TemplateFile);
                foreach (var kv in BlockEngine.InitialValues(doc, BlockEngine.ReadValues(file, doc)))
                    all.TryAdd(kv.Key, kv.Value);
            }
            catch (Exception ex)
            {
                AppLog.Error($"블록 처음 값 읽기 실패: {doc.Name}", ex);
                notes?.Add($"{doc.Name}: 어제 값을 읽지 못했습니다 ({ex.Message})");
            }
        }
        return all;
    }

    /// <summary>폰 화면 양식을 만든다.</summary>
    /// <param name="initial">블록 처음 값 (InitialBlocks)</param>
    /// <param name="isDue">문서가 그 날 만들 차례인지</param>
    /// <param name="slots">사진대지 칸 (문서, 시트) → 칸 번호 목록. 서식을 열어야 알 수 있어 밖에서 준다</param>
    public static PhoneForm BuildForm(AppConfig cfg, DateTime date, IReadOnlyDictionary<string, BlockValue> initial,
                                      Func<DocumentConfig, bool> isDue,
                                      Func<DocumentConfig, IReadOnlyList<int>> slots,
                                      string machine, DateTime now)
    {
        return new PhoneForm
        {
            Date      = DateText(date),
            UpdatedAt = now.ToString("o", CultureInfo.InvariantCulture),
            UpdatedBy = machine,
            Docs      = cfg.Documents.Select(d => new PhoneFormDoc(d.Id, d.Name, isDue(d))).ToList(),
            Blocks    = FormBlocks(cfg, initial),
            People    = FormPeople(cfg, initial),
            DocFields = FormDocFields(cfg),
            Slots     = FormSlots(cfg, slots),
        };
    }

    /// <summary>폰에서 받는 블록 종류 (날짜·계산·고정값은 PC 가 알아서 채운다)</summary>
    public static bool IsPhoneInput(FieldConfig f) => f.Type switch
    {
        BlockTypes.Value => (f.ValueSource ?? BlockTypes.SourceInput) == BlockTypes.SourceInput,
        BlockTypes.Lines or BlockTypes.Counter or BlockTypes.Table or BlockTypes.Roster or BlockTypes.Photos => true,
        _ => false,
    };

    private static List<PhoneFormBlock> FormBlocks(AppConfig cfg, IReadOnlyDictionary<string, BlockValue> initial)
    {
        var byKey = new Dictionary<string, (FieldConfig Field, List<string> Docs)>();
        foreach (var doc in cfg.Documents)
            foreach (var f in doc.Fields.Where(IsPhoneInput))
            {
                string key = BlockTypes.KeyOf(f);
                if (!byKey.TryGetValue(key, out var e)) byKey[key] = e = (f, []);
                if (!e.Docs.Contains(doc.Id)) e.Docs.Add(doc.Id);
            }

        return byKey.Select(kv =>
        {
            var (f, docs) = kv.Value;
            var v = initial.TryGetValue(kv.Key, out var found) ? found : new BlockValue();
            return new PhoneFormBlock
            {
                Key   = kv.Key,
                Label = f.Label,
                Type  = f.Type,
                Docs  = docs,
                Text  = v.Text,
                Lines = v.Lines.ToList(),
                Rows  = v.Rows.Select(r => new PhoneFormRow(r.Name, r.Spec, r.LastToday)).ToList(),
                TodayFromRoster = !string.IsNullOrWhiteSpace(f.RosterKey),
                PhotoCount = f.PhotoAreas?.Count ?? 0,
            };
        }).ToList();
    }

    /// <summary>사람 목록 = 설정의 사람 + 어제 명단에 있던 사람 (어제 온 사람은 표시)</summary>
    private static List<PhoneFormPerson> FormPeople(AppConfig cfg, IReadOnlyDictionary<string, BlockValue> initial)
    {
        var yesterday = initial.Values.SelectMany(v => v.People).ToList();
        var names     = yesterday.Select(p => p.Person.Trim()).ToHashSet();
        var list = cfg.People.Where(p => p.Name.Trim().Length > 0)
                      .Select(p => new PhoneFormPerson(p.Name.Trim(), p.Role, p.Task, names.Contains(p.Name.Trim())))
                      .ToList();
        foreach (var p in yesterday.Where(p => p.Person.Trim().Length > 0))
            if (list.All(x => x.Name != p.Person.Trim()))
                list.Add(new PhoneFormPerson(p.Person.Trim(), p.Job, p.Task, true));
        return list;
    }

    private static List<PhoneFormDocField> FormDocFields(AppConfig cfg)
    {
        var list = new List<PhoneFormDocField>();
        foreach (var doc in cfg.Documents)
            foreach (string sheet in doc.TargetSheets())
                foreach (var f in doc.Fields.Where(f => FieldTypes.IsDailyInput(f.Type)))
                {
                    string kind = f.Type == FieldTypes.Photo ? "photo" : f.Type == FieldTypes.ShapePosition ? "shape" : "input";
                    bool hasValue = kind == "input" || (kind == "photo" && !string.IsNullOrWhiteSpace(f.ValueCell));
                    var options = kind == "shape" ? (f.Presets ?? []).Select(p => p.Name).ToList() : [];
                    list.Add(new PhoneFormDocField(doc.Id, doc.Name, sheet, f.Id, f.Label, kind, hasValue, f.Suffix ?? "", options));
                }
        return list;
    }

    private static List<PhoneFormSlot> FormSlots(AppConfig cfg, Func<DocumentConfig, IReadOnlyList<int>> slots)
    {
        var list = new List<PhoneFormSlot>();
        foreach (var doc in cfg.Documents.Where(d => !string.IsNullOrWhiteSpace(d.PhotoSheet)))
        {
            var labels = doc.PhotoSlots.ToDictionary(s => s.Index, s => s.Label);
            foreach (int i in slots(doc))
                list.Add(new PhoneFormSlot(doc.Id, doc.Name, doc.PhotoSheet, i,
                    labels.TryGetValue(i, out var l) && l.Trim().Length > 0 ? l.Trim() : $"사진 {i}"));
        }
        return list;
    }

    /// <summary>양식을 inbox 폴더에 쓴다 (같은 내용이면 건드리지 않아 드라이브 동기화를 줄인다)</summary>
    public static bool WriteForm(string inboxRoot, PhoneForm form)
    {
        Directory.CreateDirectory(inboxRoot);
        string path = Path.Combine(inboxRoot, FORM_FILE_NAME);
        string json = JsonSerializer.Serialize(form, Json);
        if (File.Exists(path) && SameIgnoringStamp(File.ReadAllText(path), form)) return false;
        File.WriteAllText(path, json);
        return true;
    }

    private static bool SameIgnoringStamp(string oldJson, PhoneForm form)
    {
        try
        {
            var old = JsonSerializer.Deserialize<PhoneForm>(oldJson, Json);
            if (old == null) return false;
            var a = new PhoneForm { Date = old.Date, Docs = old.Docs, Blocks = old.Blocks, People = old.People, DocFields = old.DocFields, Slots = old.Slots };
            var b = new PhoneForm { Date = form.Date, Docs = form.Docs, Blocks = form.Blocks, People = form.People, DocFields = form.DocFields, Slots = form.Slots };
            return JsonSerializer.Serialize(a, Json) == JsonSerializer.Serialize(b, Json);
        }
        catch (JsonException) { return false; }
    }
}
