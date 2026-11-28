namespace FormFiller.Core.Config;

/// <summary>
/// [☰ 항목 선택] 목록 규칙. 여러 줄 항목(값 이름)마다 목록이 따로 있다.
/// - 처음 여는 항목: 이름에 "작업"이 들어가면 예전 공용 목록(PresetWorkItems)을 복사해 시작, 아니면 빈 목록.
/// - 추가·삭제·마지막으로 고른 것은 AppConfig 에 남는다 (저장은 부르는 쪽이 AppConfigLoader.Save).
/// </summary>
public static class PickLists
{
    private const string WORK_WORD = "작업";

    /// <summary>이 항목의 목록 (없으면 만든다)</summary>
    public static IReadOnlyList<string> Items(AppConfig config, string key) => Get(config, key).Items;

    /// <summary>마지막으로 고른 것 중 아직 목록에 있는 것</summary>
    public static IReadOnlyList<string> LastChosen(AppConfig config, string key)
    {
        var list = Get(config, key);
        return list.LastChosen.Where(list.Items.Contains).ToList();
    }

    /// <summary>목록에 넣는다. 빈 글자·이미 있는 글자면 false.</summary>
    public static bool Add(AppConfig config, string key, string text)
    {
        string t = (text ?? "").Trim();
        var list = Get(config, key);
        if (t.Length == 0 || list.Items.Contains(t)) return false;
        list.Items.Add(t);
        return true;
    }

    /// <summary>목록과 마지막 고른 것에서 뺀다</summary>
    public static void Remove(AppConfig config, string key, string text)
    {
        var list = Get(config, key);
        list.Items.Remove(text);
        list.LastChosen.Remove(text);
    }

    /// <summary>이번에 고른 것을 기억한다 (목록에 있는 것만)</summary>
    public static void SetLastChosen(AppConfig config, string key, IEnumerable<string> chosen)
    {
        var list = Get(config, key);
        list.LastChosen = chosen.Where(list.Items.Contains).Distinct().ToList();
    }

    private static ItemPickList Get(AppConfig config, string key)
    {
        string k = (key ?? "").Trim();
        if (config.ItemPickLists.TryGetValue(k, out var list)) return list;

        list = new ItemPickList
        {
            Items = k.Contains(WORK_WORD) ? [.. config.PresetWorkItems] : [],
        };
        config.ItemPickLists[k] = list;
        return list;
    }
}
