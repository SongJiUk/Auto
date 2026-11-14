namespace FormFiller.Core.Excel;

/// <summary>
/// 같은 날짜 파일을 다시 만들 때, 덮어쓰기 전에 앞의 파일을 출력 폴더의 "백업" 폴더에 남긴다.
/// 30일 지난 백업은 지운다.
/// </summary>
public static class OutputBackup
{
    public const string FolderName = "백업";
    private const int KeepDays = 30;

    /// <summary>파일이 있으면 백업 폴더로 복사하고 백업 경로를 돌려준다. 없으면 null.</summary>
    public static string? BackupIfExists(string outputPath)
    {
        if (!File.Exists(outputPath)) return null;

        string dir = Path.Combine(Path.GetDirectoryName(outputPath)!, FolderName);
        Directory.CreateDirectory(dir);
        string dest = Path.Combine(dir,
            $"{Path.GetFileNameWithoutExtension(outputPath)}_{DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(outputPath)}");
        File.Copy(outputPath, dest, overwrite: true);

        Prune(dir);
        return dest;
    }

    private static void Prune(string dir)
    {
        try
        {
            DateTime limit = DateTime.Now.AddDays(-KeepDays);
            foreach (var f in new DirectoryInfo(dir).EnumerateFiles())
                if (f.CreationTime < limit) f.Delete();
        }
        catch { /* 정리 실패는 무시 */ }
    }
}
