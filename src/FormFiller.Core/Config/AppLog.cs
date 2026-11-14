using System.Text;

namespace FormFiller.Core.Config;

/// <summary>
/// 문제를 나중에 찾을 수 있도록 프로그램 폴더의 logs\log_PC이름.txt 에 기록을 남긴다.
/// 회사·집 PC가 같은 드라이브 폴더를 쓰므로 PC마다 파일을 나눠 서로 덮어쓰지 않게 한다.
/// 기록하다 실패해도 프로그램은 멈추지 않는다.
/// </summary>
public static class AppLog
{
    private const long MaxBytes = 1_000_000;   // 1MB 넘으면 .old 로 돌리고 새로 시작
    private static readonly object Gate = new();

    public static string LogPath => Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "logs", $"log_{Environment.MachineName}.txt");

    public static void Info(string message) => Write("정보", message, null);

    public static void Error(string message, Exception? ex = null) => Write("오류", message, ex);

    private static void Write(string level, string message, Exception? ex)
    {
        try
        {
            var sb = new StringBuilder()
                .Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append(" [").Append(level).Append("] ")
                .AppendLine(message);
            if (ex != null) sb.AppendLine(ex.ToString());

            lock (Gate)
            {
                string path = LogPath;
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                var fi = new FileInfo(path);
                if (fi.Exists && fi.Length > MaxBytes)
                    File.Move(path, path + ".old", overwrite: true);
                File.AppendAllText(path, sb.ToString(), Encoding.UTF8);
            }
        }
        catch { /* 기록 실패는 무시 */ }
    }
}
