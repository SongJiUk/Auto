using System.Globalization;
using System.Text.RegularExpressions;

namespace FormFiller.Core.Config;

/// <summary>
/// 문서 주기 계산: "오늘 만들 문서인지", "다음은 언제인지", "마지막으로 언제 만들었는지".
/// 화면(App)과 떨어진 순수 계산이라 테스트할 수 있다.
///
/// [주기]
/// - daily    : 매일
/// - weekly   : 매주 WeeklyDay 요일. 그 주(월~일)에 아직 안 만들었고, 오늘이 그 요일이거나 지났으면 만든다.
/// - biweekly : 마지막으로 만든 날부터 14일 지나면
/// - monthly  : 마지막으로 만든 달이 지나면
/// - custom   : 마지막으로 만든 날부터 FrequencyDays 일 지나면
/// </summary>
public static class DocumentSchedule
{
    public const string Daily = "daily", Weekly = "weekly", Biweekly = "biweekly", Monthly = "monthly", Custom = "custom";

    private static readonly string[] KoreanDays = ["일", "월", "화", "수", "목", "금", "토"];

    /// <summary>요일 한 글자 (예: Monday → "월")</summary>
    public static string KoreanDay(DayOfWeek d) => KoreanDays[(int)d];

    /// <summary>화면에 보일 주기 이름 (예: "매일", "매주 월요일", "10일마다")</summary>
    public static string Label(DocumentConfig doc) => doc.Frequency switch
    {
        Daily    => "매일",
        Weekly   => $"매주 {KoreanDay(doc.WeeklyDay)}요일",
        Biweekly => "2주마다",
        Monthly  => "매월",
        Custom   => $"{doc.FrequencyDays}일마다",
        _        => "",
    };

    /// <summary>매일 업무인지 (아니면 매주·매월 등 "주기 업무")</summary>
    public static bool IsDaily(DocumentConfig doc) => doc.Frequency is Daily or "" or null;

    /// <summary>date 가 들어 있는 주의 월요일</summary>
    public static DateTime WeekStart(DateTime date)
    {
        int diff = ((int)date.DayOfWeek + 6) % 7;   // 월=0, 화=1, … 일=6
        return date.Date.AddDays(-diff);
    }

    /// <summary>date 가 들어 있는 주(월~일)의 그 요일 날짜</summary>
    public static DateTime DayInWeek(DateTime date, DayOfWeek day) =>
        WeekStart(date).AddDays(((int)day + 6) % 7);

    /// <summary>
    /// 파일에 쓸 날짜. 매주 문서이고 UseWeeklyDayDate 면 그 주의 정해진 요일 날짜, 아니면 고른 날짜 그대로.
    /// </summary>
    public static DateTime DocumentDate(DocumentConfig doc, DateTime selected) =>
        doc.Frequency == Weekly && doc.UseWeeklyDayDate ? DayInWeek(selected, doc.WeeklyDay) : selected.Date;

    /// <summary>오늘 만들 차례인지. last = 마지막으로 만든 날짜 (없으면 null)</summary>
    public static bool IsDue(DocumentConfig doc, DateTime today, DateTime? last)
    {
        today = today.Date;
        switch (doc.Frequency)
        {
            case Weekly:
                // 이번 주에 이미 만들었으면 끝. 아직이면 정한 요일부터 (지나도 계속) 만들 차례
                if (last != null && last.Value.Date >= WeekStart(today)) return false;
                return today >= DayInWeek(today, doc.WeeklyDay);
            case Biweekly:
                return last == null || (today - last.Value.Date).Days >= 14;
            case Monthly:
                return last == null || last.Value.Year != today.Year || last.Value.Month != today.Month;
            case Custom:
                return last == null || (today - last.Value.Date).Days >= Math.Max(1, doc.FrequencyDays);
            default:
                return true;   // daily
        }
    }

    /// <summary>다음에 만들 날짜 (오늘이 차례면 오늘)</summary>
    public static DateTime NextDue(DocumentConfig doc, DateTime today, DateTime? last)
    {
        today = today.Date;
        if (IsDue(doc, today, last)) return today;
        return doc.Frequency switch
        {
            Weekly   => last != null && last.Value.Date >= WeekStart(today)
                            ? DayInWeek(today.AddDays(7), doc.WeeklyDay)   // 이번 주 끝남 → 다음 주
                            : DayInWeek(today, doc.WeeklyDay),             // 이번 주 요일이 아직
            Biweekly => last!.Value.Date.AddDays(14),
            Monthly  => new DateTime(today.Year, today.Month, 1).AddMonths(1),
            Custom   => last!.Value.Date.AddDays(Math.Max(1, doc.FrequencyDays)),
            _        => today,
        };
    }

    /// <summary>
    /// 이번 주기 상태 한 줄 (주기 업무 목록용). 예: "이번 주 완료 (11/16)", "이번 주 아직 — 월요일부터", "다음: 12/01"
    /// </summary>
    public static string StatusText(DocumentConfig doc, DateTime today, DateTime? last)
    {
        if (IsDaily(doc)) return "";
        bool due = IsDue(doc, today, last);
        if (doc.Frequency == Weekly)
        {
            if (last != null && last.Value.Date >= WeekStart(today)) return $"이번 주 완료 ({last:MM'/'dd})";
            return due ? "이번 주 아직 — 오늘 만들 차례" : $"이번 주 {KoreanDay(doc.WeeklyDay)}요일({NextDue(doc, today, last):MM'/'dd})부터";
        }
        return due ? "오늘 만들 차례" : $"다음: {NextDue(doc, today, last):MM'/'dd}";
    }

    // ── 마지막으로 만든 날짜 찾기 ───────────────────────────────────

    /// <summary>
    /// 저장 이름 패턴을 파일 이름 찾기 식으로 바꾼다.
    /// 예) "{yy}.{MM}.{dd} 공사일보(참마루건설).xlsx" → ^(\d{2})\.(\d{2})\.(\d{2}) 공사일보\(참마루건설\)\.xlsx$
    /// </summary>
    public static Regex PatternRegex(string outputPattern)
    {
        string escaped = Regex.Escape(Path.GetFileName(outputPattern))
            .Replace(@"\{yy}", @"(?<yy>\d{2})")
            .Replace(@"\{MM}", @"(?<MM>\d{2})")
            .Replace(@"\{dd}", @"(?<dd>\d{2})");
        return new Regex("^" + escaped + "$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    }

    /// <summary>파일 이름이 패턴에 맞으면 그 날짜. 패턴에 연·월·일이 다 없거나 날짜가 틀리면 null.</summary>
    public static DateTime? DateFromFileName(string outputPattern, string fileName)
    {
        var m = PatternRegex(outputPattern).Match(fileName);
        if (!m.Success || !m.Groups["yy"].Success || !m.Groups["MM"].Success || !m.Groups["dd"].Success) return null;
        return DateTime.TryParseExact($"{m.Groups["yy"].Value}.{m.Groups["MM"].Value}.{m.Groups["dd"].Value}", "yy.MM.dd",
                                      CultureInfo.InvariantCulture, DateTimeStyles.None, out var d)
            ? d : null;
    }

    /// <summary>
    /// 이름 목록 중 패턴에 맞는 가장 최근 날짜. notAfter 를 주면 그 날짜 이후 파일은 뺀다
    /// (지난 날짜로 돌아가 볼 때 미래 파일 때문에 "이미 만듦"이 되지 않게).
    /// </summary>
    public static DateTime? LastGenerated(string outputPattern, IEnumerable<string> fileNames, DateTime? notAfter = null)
    {
        DateTime? latest = null;
        foreach (var name in fileNames)
        {
            var d = DateFromFileName(outputPattern, name);
            if (d == null || (notAfter != null && d > notAfter.Value.Date)) continue;
            if (latest == null || d > latest) latest = d;
        }
        return latest;
    }

    /// <summary>출력 폴더에서 이 문서의 가장 최근 파일 날짜</summary>
    public static DateTime? LastGenerated(DocumentConfig doc, string outputDir, DateTime? notAfter = null)
    {
        // 패턴에 하위 폴더가 있으면 (예: "공사일보\{yy}.{MM}.{dd}.xlsx") 그 폴더를 본다
        string dir = Path.Combine(outputDir, Path.GetDirectoryName(doc.OutputPattern) ?? "");
        if (!Directory.Exists(dir)) return null;
        return LastGenerated(doc.OutputPattern, Directory.EnumerateFiles(dir).Select(Path.GetFileName)!, notAfter);
    }
}
