using System.Diagnostics;

namespace FormFiller.Core.Config;

/// <summary>
/// 구글 드라이브 데스크톱(Google Drive for desktop) 관련 도우미.
/// 프로그램은 드라이브에 직접 로그인하지 않는다. PC 에 설치된 드라이브 앱이 "G:\내 드라이브" 같은 폴더로
/// 동기화해 주고, 프로그램은 그 폴더를 일반 폴더처럼 읽고 쓴다.
/// 두 PC 의 드라이브 문자가 다를 수 있어서(회사 G:, 집 H: 등), 저장된 경로가 없으면 "내 드라이브" 뒤쪽 경로로 다시 찾는다.
/// </summary>
public static class GoogleDrive
{
    /// <summary>드라이브 앱이 만드는 최상위 폴더 이름 (한글 윈도우 / 영문 윈도우)</summary>
    public static readonly string[] RootFolderNames = ["내 드라이브", "My Drive"];

    /// <summary>드라이브 앱 프로세스 이름 (작업 관리자에 보이는 GoogleDriveFS.exe)</summary>
    public const string ProcessName = "GoogleDriveFS";

    /// <summary>Google Drive 데스크톱 내려받기 페이지</summary>
    public const string DownloadUrl = "https://www.google.com/drive/download/";

    /// <summary>드라이브 앱이 실행 중인지</summary>
    public static bool IsAppRunning()
    {
        try { return Process.GetProcessesByName(ProcessName).Length > 0; }
        catch { return false; }
    }

    // FindRoots 는 드라이브 문자를 모두 훑어서 조금 느릴 수 있다. 잠깐 동안은 결과를 다시 쓴다.
    private static readonly object RootsGate = new();
    private static List<string>? _cachedRoots;
    private static DateTime _cachedAt;

    /// <summary>
    /// 이 PC 의 "내 드라이브" 폴더 목록 (예: ["G:\내 드라이브"]).
    /// 드라이브 문자 방식(G:\내 드라이브)과 폴더 미러링 방식(C:\Users\이름\My Drive)을 모두 찾는다.
    /// </summary>
    public static List<string> FindRoots(bool useCache = true)
    {
        lock (RootsGate)
        {
            if (useCache && _cachedRoots != null && DateTime.UtcNow - _cachedAt < TimeSpan.FromSeconds(30))
                return [.. _cachedRoots];
        }

        var found = new List<string>();
        try
        {
            foreach (var d in DriveInfo.GetDrives())
            {
                // CD 드라이브는 준비 확인만으로도 오래 걸릴 수 있어 건너뛴다
                if (d.DriveType == DriveType.CDRom) continue;
                try
                {
                    if (!d.IsReady) continue;
                    foreach (var name in RootFolderNames)
                    {
                        string p = Path.Combine(d.RootDirectory.FullName, name);
                        if (Directory.Exists(p)) found.Add(p);
                    }
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }

        string profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        if (!string.IsNullOrEmpty(profile))
            foreach (var name in RootFolderNames)
            {
                string p = Path.Combine(profile, name);
                if (Directory.Exists(p) && !found.Contains(p, StringComparer.OrdinalIgnoreCase)) found.Add(p);
            }

        lock (RootsGate)
        {
            _cachedRoots = found;
            _cachedAt    = DateTime.UtcNow;
        }
        return [.. found];
    }

    // ── 경로 계산 (파일을 열지 않는 순수 함수. 테스트 대상) ─────────────

    private static readonly char[] Separators = ['\\', '/'];

    /// <summary>
    /// 경로에서 "내 드라이브" 뒤쪽 부분을 돌려준다. 드라이브 안 경로가 아니면 null.
    /// 예) "G:\내 드라이브\FormFiller\사진" → "FormFiller\사진",  "G:\내 드라이브" → ""
    /// </summary>
    public static string? RelativeToDrive(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return null;
        var parts = path.Split(Separators, StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i < parts.Length; i++)
            if (RootFolderNames.Any(n => string.Equals(parts[i], n, StringComparison.OrdinalIgnoreCase)))
                return string.Join('\\', parts.Skip(i + 1));
        return null;
    }

    /// <summary>경로가 구글 드라이브 폴더 안인지 ("내 드라이브" 폴더가 경로에 들어 있는지)</summary>
    public static bool IsInsideDrive(string? path) => RelativeToDrive(path) != null;

    /// <summary>
    /// 다른 PC 에서 저장한 경로를 이 PC 경로로 바꾼다.
    /// 경로가 이미 있으면 그대로, 없으면 "내 드라이브" 뒤쪽을 이 PC 의 드라이브 폴더들에 붙여 보고 있는 것을 돌려준다.
    /// 어느 것도 없으면 원래 경로를 그대로 돌려준다 (나중에 "폴더 없음"으로 알린다).
    /// </summary>
    /// <param name="exists">폴더가 있는지 확인하는 함수 (테스트에서 바꿔 끼우려고 받는다)</param>
    public static string Remap(string path, IEnumerable<string> roots, Func<string, bool> exists)
    {
        if (string.IsNullOrWhiteSpace(path) || exists(path)) return path;
        string? rel = RelativeToDrive(path);
        if (rel == null) return path;

        foreach (var root in roots)
        {
            string candidate = JoinLike(root, rel);
            if (exists(candidate)) return candidate;
        }
        return path;
    }

    /// <summary>이 PC 의 실제 폴더로 Remap. 경로가 있으면 드라이브를 뒤지지 않는다.</summary>
    public static string Remap(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path) || Directory.Exists(path)) return path;
        return Remap(path, FindRoots(), Directory.Exists);
    }

    /// <summary>root 에 쓰인 구분자(\ 또는 /)로 rel 을 이어 붙인다.</summary>
    private static string JoinLike(string root, string rel)
    {
        if (rel.Length == 0) return root;
        char sep = root.Contains('\\') ? '\\' : Path.DirectorySeparatorChar;
        return root.TrimEnd(Separators) + sep + rel.Replace('\\', sep);
    }
}
