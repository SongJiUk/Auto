namespace FormFiller.Core.Phone;

/// <summary>칸에 넣은 사진 (TargetIndex = 넘겨받은 칸 목록의 순서)과 넣을 칸을 못 찾은 사진</summary>
public sealed record PhotoMatchResult(
    IReadOnlyList<(int TargetIndex, PhonePhoto Photo)> Placed,
    IReadOnlyList<PhonePhoto> Unplaced);

/// <summary>
/// 폰 사진의 라벨(예: "장비")로 PC 사진 칸(예: "장비 사진")을 찾는다. 띄어쓰기·대소문자는 무시한다.
/// 1) 이름이 똑같은 칸을 먼저, 2) 없으면 칸 이름에 라벨이 들어간 칸. 한 칸에는 사진 한 장만.
/// 같은 라벨 사진이 여러 장이면 같은 이름의 다음 빈 칸으로 간다 (보낸 순서 = 칸 순서).
/// </summary>
public static class PhonePhotoMatcher
{
    public static PhotoMatchResult Match(IReadOnlyList<string> targetLabels, IReadOnlyList<PhonePhoto> photos)
    {
        var keys     = targetLabels.Select(NormKey).ToList();
        var used     = new bool[keys.Count];
        var placed   = new List<(int, PhonePhoto)>();
        var unplaced = new List<PhonePhoto>();

        foreach (var photo in photos)
        {
            string label = NormKey(photo.Label);
            int index = label.Length == 0 ? -1 : FindFree(keys, used, k => k == label);
            if (index < 0 && label.Length > 0) index = FindFree(keys, used, k => k.Contains(label));

            if (index < 0) { unplaced.Add(photo); continue; }
            used[index] = true;
            placed.Add((index, photo));
        }
        return new PhotoMatchResult(placed, unplaced);
    }

    private static int FindFree(List<string> keys, bool[] used, Func<string, bool> matches)
    {
        for (int i = 0; i < keys.Count; i++)
            if (!used[i] && matches(keys[i])) return i;
        return -1;
    }

    private static string NormKey(string s) => s.Replace(" ", "").ToLowerInvariant();
}
