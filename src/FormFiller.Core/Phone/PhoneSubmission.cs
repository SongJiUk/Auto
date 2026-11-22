using System.Globalization;
using System.Text.Json;
using FormFiller.Core.Config;
using FormFiller.Core.Models;

namespace FormFiller.Core.Phone;

// ════════════════════════════════════════════════════════════════
// 폰이 올린 입력 (블록 방식, schemaVersion 2)
// inbox\yyyy-MM-dd_HHmmss_기기\input.json (+ 사진 파일)
// 바깥에서 들어온 값이라 여기서 모두 검사한다 (숫자 범위, 글자 길이, 사진 파일 이름).
// ════════════════════════════════════════════════════════════════

public sealed class SubmissionDto
{
    public int SchemaVersion { get; set; }
    public string? Date { get; set; }
    public string? Device { get; set; }
    public List<string?>? Docs { get; set; }
    public List<BlockDto?>? Blocks { get; set; }
    public List<DocFieldDto?>? DocFields { get; set; }
    public List<SlotDto?>? Slots { get; set; }
}

public sealed class BlockDto
{
    public string? Key { get; set; }
    public string? Text { get; set; }
    public List<string?>? Lines { get; set; }
    public decimal? Today { get; set; }
    public List<RowDto?>? Rows { get; set; }
    public List<PersonDto?>? People { get; set; }
    public List<PhotoFileDto?>? Photos { get; set; }
}

public sealed class RowDto { public string? Name { get; set; } public string? Spec { get; set; } public decimal Today { get; set; } }
public sealed class PersonDto { public string? Person { get; set; } public string? Job { get; set; } public string? Task { get; set; } }
public sealed class PhotoFileDto { public string? File { get; set; } public string? Caption { get; set; } }
public sealed class DocFieldDto
{
    public string? Doc { get; set; } public string? Sheet { get; set; } public string? Field { get; set; }
    public string? Value { get; set; } public string? Photo { get; set; }
}
public sealed class SlotDto { public string? Doc { get; set; } public string? Sheet { get; set; } public int Index { get; set; } public string? Photo { get; set; } }

/// <summary>검사를 마친 폰 입력</summary>
public sealed class PhoneSubmission
{
    public DateTime Date { get; init; }
    public string Device { get; init; } = "";
    public List<string> Docs { get; init; } = [];
    public Dictionary<string, BlockDto> Blocks { get; init; } = [];
    /// <summary>문서별 칸(사진·측정값·입력·위치) 값</summary>
    public List<DocFieldValue> DocFields { get; init; } = [];
    public List<(string Doc, string Sheet, int Index, string Path)> Slots { get; init; } = [];
}

/// <param name="PhotoPath">사진 전체 경로 (없으면 null)</param>
public sealed record DocFieldValue(string Doc, string Sheet, string Field, string Value, string? PhotoPath);

/// <param name="MissingPhotos">사진 파일이 아직 없음 — 드라이브 동기화 중일 수 있어 나중에 다시 본다</param>
public sealed record SubmissionReadResult(PhoneSubmission? Submission, IReadOnlyList<string> Errors, bool MissingPhotos);

public static class PhoneSubmissionReader
{
    public const string INPUT_FILE_NAME = "input.json";
    public const int MAX_COUNT = 999;
    public const int MAX_TEXT = 200;
    public const int MAX_ITEMS = 100;
    private static readonly string[] PhotoExtensions = [".jpg", ".jpeg", ".png"];

    /// <summary>input.json 의 schemaVersion (못 읽으면 0). 예전 형식(1)과 나눌 때 쓴다.</summary>
    public static int VersionOf(string folder)
    {
        try
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, INPUT_FILE_NAME)));
            return doc.RootElement.TryGetProperty("schemaVersion", out var v) && v.TryGetInt32(out int n) ? n : 0;
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException) { return 0; }
    }

    public static SubmissionReadResult Read(string folder)
    {
        string path = Path.Combine(folder, INPUT_FILE_NAME);
        SubmissionDto? dto;
        try { dto = JsonSerializer.Deserialize<SubmissionDto>(File.ReadAllText(path), PhoneDaily.Json); }
        catch (JsonException ex) { return new(null, [$"{INPUT_FILE_NAME} 형식 오류: {ex.Message}"], false); }
        catch (IOException ex)   { return new(null, [$"{INPUT_FILE_NAME} 을 읽지 못했습니다: {ex.Message}"], false); }
        if (dto == null) return new(null, [$"{INPUT_FILE_NAME} 이 비어 있습니다"], false);

        var errors = new List<string>();
        bool missing = false;
        if (dto.SchemaVersion != PhoneDaily.SCHEMA_VERSION)
            errors.Add($"모르는 형식 버전 {dto.SchemaVersion} (이 프로그램은 {PhoneDaily.SCHEMA_VERSION})");
        if (!DateTime.TryParseExact(dto.Date, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
            errors.Add($"날짜 형식이 틀렸습니다: '{dto.Date}'");

        string? PhotoPath(string? file, string what)
        {
            string f = (file ?? "").Trim();
            if (f.Length == 0) return null;
            if (f != Path.GetFileName(f) || f.Contains("..") || f.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0
                || !PhotoExtensions.Contains(Path.GetExtension(f).ToLowerInvariant()))
            { errors.Add($"{what}: 사진 파일 이름이 올바르지 않습니다 '{f}'"); return null; }
            string full = Path.Combine(folder, f);
            if (!File.Exists(full)) { missing = true; errors.Add($"{what}: 사진 파일이 아직 없습니다 '{f}'"); return null; }
            return full;
        }

        var blocks = new Dictionary<string, BlockDto>();
        foreach (var b in Take(dto.Blocks, "항목", errors))
        {
            string key = Clean(b.Key, "항목 이름", errors);
            if (key.Length == 0) continue;
            blocks[key] = new BlockDto
            {
                Key    = key,
                Text   = b.Text == null ? null : Clean(b.Text, key, errors),
                Lines  = b.Lines?.Take(MAX_ITEMS).Select(l => (string?)Clean(l, key, errors)).Where(l => l!.Length > 0).ToList(),
                Today  = b.Today is { } t ? Count(t, key, errors) : null,
                Rows   = b.Rows == null ? null : Take(b.Rows, key, errors)
                            .Select(r => new RowDto { Name = Clean(r.Name, key, errors), Spec = Clean(r.Spec, key, errors), Today = Count(r.Today, key, errors) })
                            .Where(r => r.Name!.Length > 0).Cast<RowDto?>().ToList(),
                People = b.People == null ? null : Take(b.People, key, errors)
                            .Select(p => new PersonDto { Person = Clean(p.Person, key, errors), Job = Clean(p.Job, key, errors), Task = Clean(p.Task, key, errors) })
                            .Where(p => p.Person!.Length > 0).Cast<PersonDto?>().ToList(),
                Photos = b.Photos == null ? null : Take(b.Photos, key, errors)
                            .Select(p => new PhotoFileDto { File = PhotoPath(p.File, key), Caption = Clean(p.Caption, key, errors) })
                            .Where(p => p.File != null).Cast<PhotoFileDto?>().ToList(),
            };
        }

        var docFields = new List<DocFieldValue>();
        foreach (var d in Take(dto.DocFields, "사진·측정", errors))
        {
            string doc = Clean(d.Doc, "문서", errors), sheet = Clean(d.Sheet, "시트", errors), field = Clean(d.Field, "칸", errors);
            if (doc.Length == 0 || field.Length == 0) continue;
            docFields.Add(new DocFieldValue(doc, sheet, field, Clean(d.Value, field, errors), PhotoPath(d.Photo, field)));
        }

        var slots = new List<(string, string, int, string)>();
        foreach (var s in Take(dto.Slots, "사진대지", errors))
            if (PhotoPath(s.Photo, "사진대지") is { } p)
                slots.Add((Clean(s.Doc, "문서", errors), Clean(s.Sheet, "시트", errors), s.Index, p));

        var docs = (dto.Docs ?? []).Select(d => (d ?? "").Trim()).Where(d => d.Length > 0).Distinct().ToList();
        if (docs.Count == 0) errors.Add("만들 문서를 고르지 않았습니다");

        if (errors.Count > 0) return new(null, errors, missing);
        return new(new PhoneSubmission
        {
            Date = date, Device = Clean(dto.Device, "기기", errors), Docs = docs,
            Blocks = blocks, DocFields = docFields, Slots = slots,
        }, [], false);
    }

    /// <summary>폰 값을 블록 처음 값(어제 누계 등) 위에 얹는다.</summary>
    /// <param name="fields">블록 이름 → 설정 (종류를 알려고)</param>
    /// <param name="photoFields">사진 칸인 (문서 Id, 칸 Id) — 측정값을 짝 키(#value)로 넣으려고</param>
    public static void ApplyTo(DailyInput input, PhoneSubmission sub, IReadOnlyDictionary<string, FieldConfig> fields,
                               ISet<(string Doc, string Field)> photoFields, List<string> notes)
    {
        foreach (var (key, b) in sub.Blocks)
        {
            if (!fields.TryGetValue(key, out var f)) { notes.Add($"설정에 없는 항목이라 건너뜀: {key}"); continue; }
            if (!input.Blocks.TryGetValue(key, out var v)) input.Blocks[key] = v = new BlockValue();

            if (b.Text != null)  v.Text  = b.Text;
            if (b.Lines != null) v.Lines = b.Lines.Select(l => l!).ToList();
            if (b.Today is { } t) v.Today = t;
            if (b.People != null)
                v.People = b.People.Select(p => new RosterEntry { Person = p!.Person!, Job = p.Job ?? "", Task = p.Task ?? "" }).ToList();
            if (b.Photos != null)
                v.Photos = b.Photos.Select(p => new PhotoValue { Path = p!.File!, Caption = p.Caption ?? "" }).ToList();
            if (b.Rows != null && f.Type == BlockTypes.Table) MergeRows(v, b.Rows!);
        }
        // 문서별 칸: 화면(FieldInputViewModel.FillInput)과 같은 키로 넣는다
        foreach (var d in sub.DocFields)
        {
            bool isPhoto = photoFields.Contains((d.Doc, d.Field));
            if (d.PhotoPath != null) input.DocValues[DailyInput.Key(d.Doc, d.Sheet, d.Field)] = d.PhotoPath;
            if (d.Value.Length > 0)
                input.DocValues[DailyInput.Key(d.Doc, d.Sheet, d.Field + (isPhoto ? DailyInput.ValueSuffixKey : ""))] = d.Value;
        }
    }

    /// <summary>
    /// 표: 어제 줄(전일 누계)은 그대로 두고 금일만 폰 값으로. 폰에 없는 어제 줄은 금일 0 (누계는 이어짐).
    /// 폰에서 새로 넣은 이름은 줄을 더한다.
    /// </summary>
    private static void MergeRows(BlockValue v, List<RowDto?> rows)
    {
        var phone = rows.Where(r => r != null).GroupBy(r => Norm(r!.Name)).ToDictionary(g => g.Key, g => g.First()!);
        foreach (var r in v.Rows)
            r.Today = phone.Remove(Norm(r.Name), out var p) ? p.Today : 0;
        foreach (var p in phone.Values)
            v.Rows.Add(new TableRowValue { Name = p.Name!, Spec = p.Spec ?? "", Today = p.Today });
    }

    private static string Norm(string? s) => (s ?? "").Replace(" ", "").ToUpperInvariant();

    private static IEnumerable<T> Take<T>(List<T?>? list, string what, List<string> errors) where T : class
    {
        if (list == null) return [];
        if (list.Count > MAX_ITEMS) errors.Add($"{what}: 줄이 너무 많습니다 ({list.Count}, 최대 {MAX_ITEMS})");
        return list.Take(MAX_ITEMS).Where(x => x != null).Select(x => x!);
    }

    private static string Clean(string? s, string what, List<string> errors)
    {
        string v = (s ?? "").Trim();
        if (v.Length > MAX_TEXT) { errors.Add($"{what}: 글자가 너무 깁니다 ({v.Length}자, 최대 {MAX_TEXT})"); v = v[..MAX_TEXT]; }
        return v;
    }

    private static decimal Count(decimal n, string what, List<string> errors)
    {
        if (n < 0 || n > MAX_COUNT) { errors.Add($"{what}: 숫자 {n} 가 0~{MAX_COUNT} 범위 밖입니다"); return 0; }
        return n;
    }
}
