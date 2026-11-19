using System.Globalization;

namespace FormFiller.Core.Config;

/// <summary>
/// 블록(범용 항목) 종류와 규칙 이름.
/// "이 칸이 무슨 뜻인지"를 코드에 박지 않고, 문서를 등록할 때 사용자가 종류·규칙을 고른다.
///
///   value   글자·숫자 하나        — 매일 입력 / 어제 값 그대로 / 고정값
///   date    날짜                  — 형식을 고른다 ({yyyy}.{MM}.{dd} 등)
///   lines   여러 줄 글            — 앞 글자("- "), 처음 값(어제 이 블록 / 어제 다른 블록)
///   counter 숫자 누계             — 오늘 칸 + (전일 칸) + 누계 칸. 누계 = 어제 누계 + 오늘
///   table   표 (이름 줄마다 숫자)  — 열마다 역할: 이름·규격·전일·금일·누계. 금일을 명단에서 셀 수도 있다
///   roster  명단 (사람 한 줄씩)    — 열마다 역할: 직종·성명·작업내용. 사람 목록에서 오늘 나온 사람을 고른다
///   calc    계산                  — 다른 블록의 합계 / 이름 나열 / 명단 인원 수
///   photos  사진 구역             — 사진 칸들(+설명 칸). 그날 장수만큼 채우고, 1장인 날 자리는 따로 정할 수 있다
/// </summary>
public static class BlockTypes
{
    public const string Value = "value", Date = "date", Lines = "lines", Counter = "counter",
                        Table = "table", Roster = "roster", Calc = "calc", Photos = "photos", Formula = "formula";

    public static readonly string[] All = [Value, Date, Lines, Counter, Table, Roster, Calc, Photos, Formula];

    // ── 0 일 때 표시 ───────────────────────────────────────────────────
    public const string ZeroBlank = "blank", ZeroDash = "dash";

    /// <summary>0 대신 쓸 글자 (null 이면 0 그대로)</summary>
    public static string? ZeroText(string? zeroAs) => zeroAs switch
    {
        ZeroBlank => "",
        ZeroDash  => "-",
        _         => null,
    };

    public static bool IsBlock(string type) => All.Contains(type);

    /// <summary>블록의 값 이름 (Key 가 비어 있으면 Label)</summary>
    public static string KeyOf(FieldConfig f) => string.IsNullOrWhiteSpace(f.Key) ? f.Label.Trim() : f.Key.Trim();

    // ── value 출처 ───────────────────────────────────────────────────
    public const string SourceInput = "input", SourceKeep = "keep", SourceFixed = "fixed";

    // ── 표·명단 열 역할 ──────────────────────────────────────────────
    public const string RoleName = "name", RoleSpec = "spec", RolePrev = "prev", RoleToday = "today", RoleTotal = "total";
    public const string RoleNo = "no", RoleJob = "job", RolePerson = "person", RoleTask = "task";

    /// <summary>표: 고를 열 (순서대로 안내). 필수 = 이름·금일</summary>
    public static readonly (string Role, string Title, bool Required)[] TableRoles =
    [
        (RoleName, "이름", true), (RoleSpec, "규격", false), (RolePrev, "전일", false),
        (RoleToday, "금일", true), (RoleTotal, "누계", false),
    ];

    /// <summary>명단: 고를 열. 필수 = 성명</summary>
    public static readonly (string Role, string Title, bool Required)[] RosterRoles =
    [
        (RoleJob, "직종", false), (RolePerson, "성명", true), (RoleTask, "작업내용", false),
    ];

    public static string RoleTitle(string role) =>
        TableRoles.Concat(RosterRoles).FirstOrDefault(r => r.Role == role).Title ?? role;

    public static string? ColOf(FieldConfig f, string role) =>
        f.Columns?.FirstOrDefault(c => c.Role == role)?.Col is { Length: > 0 } c ? c : null;

    // ── 계산 ─────────────────────────────────────────────────────────
    public const string CalcSum = "sum", CalcNames = "names", CalcCount = "count";

    // ── 날짜 형식 ────────────────────────────────────────────────────

    /// <summary>설정 화면에서 고르는 날짜 형식 (직접 입력도 가능)</summary>
    public static readonly string[] DateFormats =
    [
        "{yyyy}년 {MM}월 {dd}일 {요일}",
        "{yyyy}.{MM}.{dd}",
        "{yyyy}. {MM}. {dd}",
        "{yy}.{MM}.{dd}",
        "{yyyy}-{MM}-{dd}",
        "{yyyy}년 {M}월 {d}일",
        "{yyyy}.{MM}.{dd}({요})",
        "{M}/{d}",
    ];

    private static readonly string[] Days = ["일", "월", "화", "수", "목", "금", "토"];

    /// <summary>날짜를 형식대로 글자로 (예: "{yyyy}.{MM}.{dd}({요})" → "2026.11.18(수)")</summary>
    public static string FormatDate(string? format, DateTime d)
    {
        format = string.IsNullOrWhiteSpace(format) ? DateFormats[0] : format;
        string day = Days[(int)d.DayOfWeek];
        return format
            .Replace("{yyyy}", d.Year.ToString("0000", CultureInfo.InvariantCulture))
            .Replace("{yy}",   (d.Year % 100).ToString("00", CultureInfo.InvariantCulture))
            .Replace("{MM}",   d.Month.ToString("00", CultureInfo.InvariantCulture))
            .Replace("{M}",    d.Month.ToString(CultureInfo.InvariantCulture))
            .Replace("{dd}",   d.Day.ToString("00", CultureInfo.InvariantCulture))
            .Replace("{d}",    d.Day.ToString(CultureInfo.InvariantCulture))
            .Replace("{요일}", day + "요일")
            .Replace("{요}",   day);
    }

    /// <summary>"12", "3.5", "1,234" 같은 글자를 숫자로. 숫자가 아니면 null.</summary>
    public static decimal? ParseNumber(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        return decimal.TryParse(s.Trim().Replace(",", ""), NumberStyles.Number, CultureInfo.InvariantCulture, out var v) ? v : null;
    }
}
