using System.Globalization;

namespace FormFiller.Core.Config;

/// <summary>공유 표시 띠 한 줄 + (있으면) 다른 PC 변경 알림</summary>
public sealed record SharedConfigStatus(bool IsShared, string Text, string? Notice);

/// <summary>
/// 모든 PC 가 같은 설정(구글 드라이브의 config.json)을 쓰는지 보여 주고,
/// 다른 PC 에서 설정을 바꿨으면 켤 때 알린다 (2026-11-23 화면 정리 4단계).
/// "이 PC 가 마지막으로 본 저장 시각"은 PC 마다 다르므로 드라이브가 아니라 PC 안(LocalAppData)에 둔다.
/// </summary>
public static class SharedConfig
{
    private const string TimeFormat = "MM'/'dd HH:mm";

    /// <param name="baseDir">exe 폴더 (config.json 위치)</param>
    /// <param name="machine">이 PC 이름</param>
    /// <param name="lastSeen">이 PC 가 지난번에 확인한 설정 저장 시각 (처음 켠 PC 면 null)</param>
    public static SharedConfigStatus Describe(AppConfig cfg, string baseDir, string machine, DateTime? lastSeen)
    {
        string? rel = GoogleDrive.RelativeToDrive(baseDir);
        if (rel == null)
            return new(false, $"⚠ 이 PC 에만 있는 설정입니다 — 다른 PC·폰과 공유되지 않습니다 ({baseDir.TrimEnd('\\')})", null);

        string folder = rel.Length == 0 ? "내 드라이브" : $@"내 드라이브\{rel.TrimEnd('\\')}";
        string saved  = cfg.LastSavedAt is { } at
            ? $"마지막 저장: {Who(cfg.LastSavedBy, machine)} {Fmt(at)}"
            : "마지막 저장 기록 없음";

        string? notice = null;
        bool byOther = !string.IsNullOrWhiteSpace(cfg.LastSavedBy)
                       && !string.Equals(cfg.LastSavedBy, machine, StringComparison.OrdinalIgnoreCase);
        if (byOther && lastSeen is { } seen && cfg.LastSavedAt is { } changed && changed > seen)
            notice = $"🔔 다른 PC({cfg.LastSavedBy})에서 {Fmt(changed)}에 설정을 바꿨습니다";

        return new(true, $"☁ 공유 설정 · {folder} · {saved}", notice);
    }

    /// <summary>저장할 때 누가·언제 저장했는지 남긴다</summary>
    public static void Stamp(AppConfig cfg, string machine, DateTime now)
    {
        cfg.LastSavedBy = machine;
        cfg.LastSavedAt = now;
    }

    /// <summary>이 PC 가 마지막으로 확인한 저장 시각. 파일이 없거나 깨졌으면 null.</summary>
    public static DateTime? ReadSeen(string path)
    {
        try
        {
            if (!File.Exists(path)) return null;
            return DateTime.TryParse(File.ReadAllText(path).Trim(), CultureInfo.InvariantCulture,
                                     DateTimeStyles.RoundtripKind, out var d) ? d : null;
        }
        catch (IOException) { return null; }
        catch (UnauthorizedAccessException) { return null; }
    }

    public static void WriteSeen(string path, DateTime seen)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, seen.ToString("o", CultureInfo.InvariantCulture));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            AppLog.Error("설정 확인 시각 기록 실패", ex);
        }
    }

    /// <summary>이 PC 의 "본 시각" 파일 (드라이브와 무관한 PC 안 폴더)</summary>
    public static string DefaultSeenPath() =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "업무자동화", "config_seen.txt");

    private static string Who(string by, string machine) =>
        string.IsNullOrWhiteSpace(by) ? "알 수 없는 PC"
        : string.Equals(by, machine, StringComparison.OrdinalIgnoreCase) ? "이 PC" : by;

    private static string Fmt(DateTime d) => d.ToString(TimeFormat, CultureInfo.InvariantCulture);
}
