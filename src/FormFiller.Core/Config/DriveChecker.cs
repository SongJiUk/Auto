namespace FormFiller.Core.Config;

/// <summary>점검 결과 단계</summary>
public enum CheckState { Ok, Warning, Error }

/// <summary>
/// 점검 한 줄. 화면(구글 드라이브 창)이 그대로 보여 준다.
/// </summary>
/// <param name="Title">무엇을 점검했는지 (예: "Google Drive 앱")</param>
/// <param name="State">결과</param>
/// <param name="Detail">결과 설명 (경로 등)</param>
/// <param name="Fix">해결 방법 (문제가 있을 때만)</param>
/// <param name="Important">시작할 때 상단 띠로 알릴 만큼 중요한지</param>
/// <param name="OfferDownload">[다운로드 페이지 열기] 버튼을 보일지</param>
public record DriveCheckItem(
    string Title, CheckState State, string Detail,
    string? Fix = null, bool Important = false, bool OfferDownload = false);

/// <summary>
/// 구글 드라이브로 두 PC 가 같은 프로그램·설정·결과를 쓰는지 점검한다.
/// 파일을 조금 건드리므로(쓰기 시험) 화면 스레드 밖에서 부른다.
/// </summary>
public static class DriveChecker
{
    public static List<DriveCheckItem> Run(AppConfig config, string baseDir)
    {
        var items = new List<DriveCheckItem>();
        var roots = GoogleDrive.FindRoots(useCache: false);

        // 1. 드라이브 앱
        items.Add(GoogleDrive.IsAppRunning()
            ? new("Google Drive 앱", CheckState.Ok, "실행 중")
            : new("Google Drive 앱", CheckState.Error,
                  "Google Drive 데스크톱이 꺼져 있거나 설치되어 있지 않습니다. 꺼져 있으면 다른 PC 와 동기화되지 않습니다.",
                  "설치되어 있으면 시작 메뉴에서 'Google Drive' 를 실행하세요. 없으면 내려받아 설치하고 같은 구글 계정으로 로그인하세요.",
                  Important: true, OfferDownload: true));

        // 2. 내 드라이브 폴더
        items.Add(roots.Count > 0
            ? new("드라이브 폴더", CheckState.Ok, string.Join(", ", roots))
            : new("드라이브 폴더", CheckState.Error,
                  "이 PC 에서 '내 드라이브' 폴더를 찾지 못했습니다.",
                  "Google Drive 앱에 로그인했는지 확인하세요. 로그인하면 탐색기에 'Google Drive (G:)' 같은 드라이브가 생깁니다.",
                  Important: true));

        // 3. 프로그램 위치 (드라이브 밖 복사본을 실행하면 설정·결과가 다른 PC 와 따로 논다)
        items.Add(IsInside(baseDir, roots)
            ? new("프로그램 위치", CheckState.Ok, baseDir)
            : new("프로그램 위치", CheckState.Warning,
                  $"드라이브 밖에서 실행 중입니다: {baseDir}",
                  "이 복사본의 설정과 만든 파일은 다른 PC 와 동기화되지 않습니다. 드라이브 폴더 안의 프로그램(바탕화면 바로가기)을 실행하세요.",
                  Important: true));

        // 4. 출력 폴더
        items.Add(CheckOutputFolder(config, baseDir, roots));

        // 5. 문서별 사진 폴더 (사진 칸이 있는 문서만)
        foreach (var doc in config.Documents.Where(d => d.Fields.Any(f => f.Type == FieldTypes.Photo)))
            items.Add(CheckPhotoFolder(doc, roots));

        return items;
    }

    /// <summary>시작할 때 알릴 문제가 있는지</summary>
    public static bool HasImportantProblem(IEnumerable<DriveCheckItem> items) =>
        items.Any(i => i.Important && i.State != CheckState.Ok);

    /// <summary>출력 폴더의 실제 경로 (상대 경로면 프로그램 폴더 기준)</summary>
    public static string ResolveOutputFolder(AppConfig config, string baseDir)
    {
        string p = string.IsNullOrWhiteSpace(config.OutputFolder) ? "output" : config.OutputFolder;
        return Path.IsPathRooted(p) ? p : Path.GetFullPath(Path.Combine(baseDir, p));
    }

    private static DriveCheckItem CheckOutputFolder(AppConfig config, string baseDir, List<string> roots)
    {
        const string title = "출력 폴더";
        string path = ResolveOutputFolder(config, baseDir);

        if (!Directory.Exists(path))
            return new(title, CheckState.Warning, $"아직 없습니다: {path}", "처음 파일을 만들 때 생깁니다.");

        if (!CanWrite(path, out string? error))
            return new(title, CheckState.Error, $"쓸 수 없습니다: {path}\n{error}",
                       "폴더 권한을 확인하거나 다른 폴더를 고르세요.", Important: true);

        return IsInside(path, roots)
            ? new(title, CheckState.Ok, path)
            : new(title, CheckState.Warning, $"드라이브 밖 폴더입니다: {path}",
                  "다른 PC 에서 어제 파일(전일 누계)을 찾지 못합니다. 드라이브 안 폴더를 고르세요.");
    }

    private static DriveCheckItem CheckPhotoFolder(DocumentConfig doc, List<string> roots)
    {
        string title = $"사진 폴더 · {doc.Name}";
        if (string.IsNullOrWhiteSpace(doc.PhotoFolder))
            return new(title, CheckState.Warning, "아직 정하지 않았습니다.", "[찾아보기] 로 사진을 올리는 드라이브 폴더를 고르세요.");

        // 다른 PC 에서 고른 경로면 이 PC 의 드라이브 문자로 바꿔서 확인
        string path = GoogleDrive.Remap(doc.PhotoFolder, roots, Directory.Exists);
        if (!Directory.Exists(path))
            return new(title, CheckState.Error, $"폴더가 없습니다: {doc.PhotoFolder}",
                       "드라이브 앱이 켜져 있는지 확인하거나 [찾아보기] 로 다시 고르세요.");

        return IsInside(path, roots)
            ? new(title, CheckState.Ok, path)
            : new(title, CheckState.Warning, $"드라이브 밖 폴더입니다: {path}",
                  "다른 PC 에서는 이 폴더를 찾지 못합니다.");
    }

    /// <summary>roots 중 하나의 안쪽이거나, 경로에 '내 드라이브' 가 들어 있으면 드라이브 안으로 본다.</summary>
    private static bool IsInside(string path, List<string> roots)
    {
        string full;
        try { full = Path.GetFullPath(path); } catch { full = path; }
        return roots.Any(r => full.StartsWith(r, StringComparison.OrdinalIgnoreCase)) || GoogleDrive.IsInsideDrive(full);
    }

    /// <summary>폴더에 작은 파일을 만들었다 지워서 쓰기가 되는지 본다.</summary>
    private static bool CanWrite(string dir, out string? error)
    {
        string probe = Path.Combine(dir, $".write_test_{Environment.MachineName}.tmp");
        try
        {
            File.WriteAllText(probe, "");
            File.Delete(probe);
            error = null;
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            error = ex.Message;
            return false;
        }
    }
}
