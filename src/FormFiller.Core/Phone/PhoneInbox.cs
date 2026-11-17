using System.Globalization;
using System.Text.Json;

namespace FormFiller.Core.Phone;

/// <summary>
/// 드라이브 inbox 폴더에서 폰 입력을 찾고 읽는다.
/// 폴더 구조: inbox\yyyy-MM-dd_HHmmss_기기이름\input.json (+ 사진 파일)
/// 폰에서 온 파일은 바깥에서 들어온 값이라 믿지 않고 여기서 모두 검증한다 (숫자 범위, 사진 파일 이름 등).
/// </summary>
public static class PhoneInbox
{
    public const string INPUT_FILE_NAME = "input.json";
    public const int SCHEMA_VERSION = 1;
    public const int MAX_COUNT = 999;
    public const int MAX_TEXT_LENGTH = 200;
    public const string TEAM_CONSTRUCTION = "공사";
    public const string TEAM_SUJANGJAE = "수장재";
    private const string DATE_FORMAT = "yyyy-MM-dd";

    private static readonly string[] AllowedPhotoExtensions = [".jpg", ".jpeg", ".png"];
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    /// <summary>설정에 inbox 폴더가 없을 때: exe 폴더 옆 (예: G:\내 드라이브\FormFiller_inbox)</summary>
    public static readonly string DefaultInboxFolder = Path.Combine("..", "FormFiller_inbox");

    /// <summary>설정값(비었거나 exe 기준 상대 경로, 또는 절대 경로)을 실제 폴더 경로로 바꾼다.</summary>
    public static string ResolveInboxFolder(string configured, string baseDir)
    {
        string folder = string.IsNullOrWhiteSpace(configured) ? DefaultInboxFolder : configured.Trim();
        return Path.GetFullPath(Path.Combine(baseDir, folder));   // folder 가 절대 경로면 Combine 이 그대로 돌려준다
    }

    /// <summary>그 날짜의 입력 폴더 목록 (보낸 순서). input.json 이 아직 동기화 안 된 폴더는 뺀다.</summary>
    public static IReadOnlyList<string> FindSubmissions(string inboxRoot, DateTime date)
    {
        if (!Directory.Exists(inboxRoot)) return [];
        string prefix = date.ToString(DATE_FORMAT, CultureInfo.InvariantCulture) + "_";
        return Directory.GetDirectories(inboxRoot)
            .Where(d => Path.GetFileName(d).StartsWith(prefix, StringComparison.Ordinal))
            .Where(d => File.Exists(Path.Combine(d, INPUT_FILE_NAME)))
            .OrderBy(d => Path.GetFileName(d), StringComparer.Ordinal)
            .ToList();
    }

    /// <summary>입력 폴더 하나를 읽고 검증한다. 예외를 던지지 않고 Errors 로 알려 준다.</summary>
    public static PhoneReadResult Read(string submissionFolder)
    {
        string path = Path.Combine(submissionFolder, INPUT_FILE_NAME);
        if (!File.Exists(path))
            return Fail($"{INPUT_FILE_NAME} 이 없습니다 (아직 동기화 중일 수 있음)");

        PhoneInputDto? dto;
        try { dto = JsonSerializer.Deserialize<PhoneInputDto>(File.ReadAllText(path), JsonOptions); }
        catch (JsonException ex) { return Fail($"{INPUT_FILE_NAME} 형식 오류: {ex.Message}"); }
        catch (IOException ex) { return Fail($"{INPUT_FILE_NAME} 을 읽지 못했습니다: {ex.Message}"); }
        if (dto == null) return Fail($"{INPUT_FILE_NAME} 이 비어 있습니다");

        var errors = new List<string>();
        var input = Validate(dto, submissionFolder, errors);
        return errors.Count == 0 ? new PhoneReadResult(input, []) : new PhoneReadResult(null, errors);
    }

    private static PhoneReadResult Fail(string error) => new(null, [error]);

    private static PhoneInput Validate(PhoneInputDto dto, string folder, List<string> errors)
    {
        if (dto.SchemaVersion != SCHEMA_VERSION)
            errors.Add($"모르는 형식 버전 {dto.SchemaVersion} (이 프로그램은 {SCHEMA_VERSION})");

        DateTime date = default;
        if (!DateTime.TryParseExact(dto.Date, DATE_FORMAT, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            errors.Add($"날짜 형식이 틀렸습니다: '{dto.Date}' ({DATE_FORMAT})");
        else if (!Path.GetFileName(folder).StartsWith(dto.Date + "_", StringComparison.Ordinal))
            errors.Add($"폴더 이름의 날짜와 input.json 의 날짜({dto.Date})가 다릅니다");

        return new PhoneInput
        {
            Date               = date,
            Device             = Text(dto.Device, "기기 이름", errors),
            Weather            = Text(dto.Weather, "날씨", errors),
            TodayMorningWork   = Lines(dto.TodayMorningWork, "금일 오전 작업", errors),
            TodayAfternoonWork = Lines(dto.TodayAfternoonWork, "금일 오후 작업", errors),
            TomorrowWork       = Lines(dto.TomorrowWork, "명일 작업", errors),
            Equipment          = Counts(dto.Equipment, "장비", errors),
            Personnel          = Counts(dto.Personnel, "인원", errors),
            SujangjaePersonnel = Counts(dto.SujangjaePersonnel, "수장재 인원", errors),
            Attendees          = Attendees(dto.Attendees, errors),
            Photos             = Photos(dto.Photos, folder, errors),
        };
    }

    private static string Text(string? value, string what, List<string> errors)
    {
        string v = (value ?? "").Trim();
        if (v.Length > MAX_TEXT_LENGTH) errors.Add($"{what}이(가) 너무 깁니다 ({v.Length}자, 최대 {MAX_TEXT_LENGTH})");
        return v;
    }

    private static List<string> Lines(List<string?>? lines, string what, List<string> errors) =>
        (lines ?? []).Select(l => Text(l, what, errors)).Where(l => l.Length > 0).ToList();

    private static List<PhoneCount> Counts(List<CountDto?>? rows, string what, List<string> errors)
    {
        var result = new List<PhoneCount>();
        foreach (var r in rows ?? [])
        {
            string name = Text(r?.Name ?? r?.Role, what + " 이름", errors);
            if (name.Length == 0) { errors.Add($"{what}: 이름 없는 줄이 있습니다"); continue; }
            int count = r!.TodayCount;
            if (count < 0 || count > MAX_COUNT) errors.Add($"{what} '{name}': 금일 수 {count} 는 0~{MAX_COUNT} 범위 밖입니다");
            result.Add(new PhoneCount(name, Text(r.Spec, what + " 규격", errors), count));
        }
        return result;
    }

    private static List<PhoneAttendee> Attendees(List<AttendeeDto?>? rows, List<string> errors)
    {
        var result = new List<PhoneAttendee>();
        foreach (var r in rows ?? [])
        {
            string name = Text(r?.Name, "출근자 이름", errors);
            string team = Text(r?.Team, "출근자 팀", errors);
            if (name.Length == 0) { errors.Add("출근자: 이름 없는 줄이 있습니다"); continue; }
            // 이름은 개인정보라 오류 문구에 넣지 않는다 (몇 번째 줄인지만)
            if (team != TEAM_CONSTRUCTION && team != TEAM_SUJANGJAE)
                errors.Add($"출근자 {result.Count + 1}번째: 팀은 '{TEAM_CONSTRUCTION}' 또는 '{TEAM_SUJANGJAE}' 이어야 합니다");
            result.Add(new PhoneAttendee(name, Text(r!.Role, "출근자 직종", errors), team));
        }
        return result;
    }

    private static List<PhonePhoto> Photos(List<PhotoDto?>? rows, string folder, List<string> errors)
    {
        var result = new List<PhonePhoto>();
        foreach (var r in rows ?? [])
        {
            string file = (r?.File ?? "").Trim();
            // 파일 이름만 허용: 경로 문자가 들어가면 inbox 폴더 밖을 가리킬 수 있다
            if (file.Length == 0 || file != Path.GetFileName(file) || file.Contains("..")
                || file.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            { errors.Add($"사진 파일 이름이 올바르지 않습니다: '{file}'"); continue; }
            if (!AllowedPhotoExtensions.Contains(Path.GetExtension(file).ToLowerInvariant()))
            { errors.Add($"사진은 jpg/png 만 됩니다: '{file}'"); continue; }
            string full = Path.Combine(folder, file);
            if (!File.Exists(full)) { errors.Add($"사진 파일이 없습니다 (동기화 중일 수 있음): '{file}'"); continue; }
            result.Add(new PhonePhoto(file, Text(r!.Label, "사진 라벨", errors), Text(r.Value, "측정값", errors), full));
        }
        return result;
    }

    // ── JSON 그대로 받는 그릇 (검증 전이라 모두 비어 있을 수 있음) ─────────
    private sealed class PhoneInputDto
    {
        public int SchemaVersion { get; set; }
        public string? Date { get; set; }
        public string? Device { get; set; }
        public string? Weather { get; set; }
        public List<string?>? TodayMorningWork { get; set; }
        public List<string?>? TodayAfternoonWork { get; set; }
        public List<string?>? TomorrowWork { get; set; }
        public List<CountDto?>? Equipment { get; set; }
        public List<CountDto?>? Personnel { get; set; }
        public List<CountDto?>? SujangjaePersonnel { get; set; }
        public List<AttendeeDto?>? Attendees { get; set; }
        public List<PhotoDto?>? Photos { get; set; }
    }

    private sealed class CountDto
    {
        public string? Name { get; set; }
        public string? Role { get; set; }
        public string? Spec { get; set; }
        public int TodayCount { get; set; }
    }

    private sealed class AttendeeDto
    {
        public string? Name { get; set; }
        public string? Role { get; set; }
        public string? Team { get; set; }
    }

    private sealed class PhotoDto
    {
        public string? File { get; set; }
        public string? Label { get; set; }
        public string? Value { get; set; }
    }
}
