using FormFiller.Core.Models;

namespace FormFiller.Core.Phone;

/// <summary>합친 결과. Unmatched = 화면에 없어서 새로 넣은 줄, Conflicts = 화면 값과 달라서 폰 값으로 바꾼 줄</summary>
public sealed record PhoneMergeResult(DailyInput Input, IReadOnlyList<string> Unmatched, IReadOnlyList<string> Conflicts);

/// <summary>
/// 폰 입력을 지금 화면 값에 합친다. 원래 화면 값(DailyInput)은 바꾸지 않고 새 DailyInput 을 돌려준다.
/// 규칙:
/// - 장비·인원은 이름(띄어쓰기·대소문자 무시)으로 줄을 찾아 "금일"만 바꾼다. 전일·누계는 그대로.
/// - 폰이 보내지 않은 줄은 화면 값 그대로.
/// - 화면에 없는 이름은 새 줄(전일 0)로 넣고 Unmatched 로 알린다.
/// - 날씨·작업사항은 폰에 값이 있을 때만 바꾼다.
/// </summary>
public static class PhoneInputMerger
{
    public static PhoneMergeResult Apply(DailyInput screen, PhoneInput phone)
    {
        if (screen.Date.Date != phone.Date.Date)
            throw new ArgumentException($"날짜가 다릅니다: 화면 {screen.Date:yyyy-MM-dd}, 폰 {phone.Date:yyyy-MM-dd}", nameof(phone));

        var unmatched = new List<string>();
        var conflicts = new List<string>();

        var merged = new DailyInput
        {
            Date               = screen.Date,
            Weather            = phone.Weather.Length > 0 ? phone.Weather : screen.Weather,
            TodayMorningWork   = JoinOr(phone.TodayMorningWork, screen.TodayMorningWork),
            TodayAfternoonWork = JoinOr(phone.TodayAfternoonWork, screen.TodayAfternoonWork),
            TomorrowWork       = JoinOr(phone.TomorrowWork, screen.TomorrowWork),
            TbmPersonnel       = screen.TbmPersonnel,
            DocValues          = new Dictionary<string, string>(screen.DocValues),
            Equipment          = MergeEquipment(screen.Equipment, phone.Equipment, unmatched, conflicts),
            Personnel          = MergePersonnel(screen.Personnel, phone.Personnel, "인원", unmatched, conflicts),
            SujangjaePersonnel = MergePersonnel(screen.SujangjaePersonnel, phone.SujangjaePersonnel, "수장재", unmatched, conflicts),
        };
        return new PhoneMergeResult(merged, unmatched, conflicts);
    }

    private static string NormKey(string s) => s.Replace(" ", "").ToLowerInvariant();

    private static string JoinOr(IReadOnlyList<string> phoneLines, string screenText) =>
        phoneLines.Count > 0 ? string.Join("\n", phoneLines) : screenText;

    private static Dictionary<string, PhoneCount> ByName(IReadOnlyList<PhoneCount> rows) =>
        rows.GroupBy(r => NormKey(r.Name)).ToDictionary(g => g.Key, g => g.Last());

    private static int Pick(string what, string name, int screenToday, PhoneCount? phoneRow, List<string> conflicts)
    {
        if (phoneRow == null) return screenToday;
        if (screenToday != 0 && screenToday != phoneRow.TodayCount)
            conflicts.Add($"{what} '{name}': 화면 {screenToday} → 폰 {phoneRow.TodayCount}");
        return phoneRow.TodayCount;
    }

    private static List<EquipmentEntry> MergeEquipment(
        List<EquipmentEntry> screen, IReadOnlyList<PhoneCount> phone, List<string> unmatched, List<string> conflicts)
    {
        var phoneByName = ByName(phone);
        var result = screen.Select(e =>
        {
            phoneByName.Remove(NormKey(e.Name), out var p);
            return new EquipmentEntry
            {
                Name               = e.Name,
                Spec               = e.Spec,
                PrevDayCount       = e.PrevDayCount,
                LastFileTodayCount = e.LastFileTodayCount,
                TodayCount         = Pick("장비", e.Name, e.TodayCount, p, conflicts),
            };
        }).ToList();

        foreach (var p in phone.Where(p => phoneByName.ContainsKey(NormKey(p.Name))).DistinctBy(p => NormKey(p.Name)))
        {
            var last = phoneByName[NormKey(p.Name)];
            unmatched.Add($"장비 '{last.Name}' (화면에 없던 줄 — 전일 0 으로 추가)");
            result.Add(new EquipmentEntry { Name = last.Name, Spec = last.Spec, TodayCount = last.TodayCount });
        }
        return result;
    }

    private static List<PersonnelEntry> MergePersonnel(
        List<PersonnelEntry> screen, IReadOnlyList<PhoneCount> phone, string what,
        List<string> unmatched, List<string> conflicts)
    {
        var phoneByName = ByName(phone);
        var result = screen.Select(e =>
        {
            phoneByName.Remove(NormKey(e.Role), out var p);
            return new PersonnelEntry
            {
                Role                = e.Role,
                PrevCumulativeCount = e.PrevCumulativeCount,
                LastFileTodayCount  = e.LastFileTodayCount,
                TodayCount          = Pick(what, e.Role, e.TodayCount, p, conflicts),
            };
        }).ToList();

        foreach (var p in phone.Where(p => phoneByName.ContainsKey(NormKey(p.Name))).DistinctBy(p => NormKey(p.Name)))
        {
            var last = phoneByName[NormKey(p.Name)];
            unmatched.Add($"{what} '{last.Name}' (화면에 없던 줄 — 전일누계 0 으로 추가)");
            result.Add(new PersonnelEntry { Role = last.Name, TodayCount = last.TodayCount });
        }
        return result;
    }
}
