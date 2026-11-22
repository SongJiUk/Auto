using System.Globalization;
using System.Text.Json;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Core.Models;

namespace FormFiller.Core.Phone;

/// <summary>처리 결과 (폰 화면이 읽어서 보여 준다)</summary>
public sealed class PhoneResult
{
    /// <summary>done · partial · failed</summary>
    public string Status { get; init; } = "";
    public string At { get; init; } = "";
    public string By { get; init; } = "";
    public List<string> Files { get; init; } = [];
    public List<string> Messages { get; init; } = [];
}

/// <summary>
/// 폰 입력 → 문서 자동 생성 (2026-11-23 사용자 결정: 출근해서 켜 둔 PC 가 만든다).
/// 폰이 올린 폴더마다: 처음 값(어제 누계) 읽기 → 폰 값 얹기 → PC 설정대로 생성 → result.json 남기기.
/// result.json 이 있는 폴더는 다시 만들지 않는다.
/// </summary>
public static class PhoneAutoGenerator
{
    public const string RESULT_FILE_NAME = "result.json";
    public const string CLAIM_FILE_NAME  = "claim.txt";
    /// <summary>사진이 아직 안 올라온 폴더를 이만큼 기다린 뒤에는 실패로 남긴다</summary>
    public static readonly TimeSpan PhotoWait  = TimeSpan.FromMinutes(30);
    /// <summary>다른 PC 가 맡은 지 이만큼 안 됐으면 건드리지 않는다 (두 PC 가 함께 켜져 있을 때)</summary>
    public static readonly TimeSpan ClaimValid = TimeSpan.FromMinutes(10);

    /// <summary>아직 처리 안 한 폰 입력 폴더 (새 형식만, 오래된 것부터)</summary>
    public static List<string> FindPending(string inboxRoot)
    {
        if (!Directory.Exists(inboxRoot)) return [];
        return Directory.GetDirectories(inboxRoot)
            .Where(d => File.Exists(Path.Combine(d, PhoneSubmissionReader.INPUT_FILE_NAME)))
            .Where(d => !File.Exists(Path.Combine(d, RESULT_FILE_NAME)))
            .Where(d => PhoneSubmissionReader.VersionOf(d) == PhoneDaily.SCHEMA_VERSION)
            .OrderBy(d => Path.GetFileName(d), StringComparer.Ordinal)
            .ToList();
    }

    /// <summary>
    /// 폴더 하나를 처리한다. 사진 동기화를 기다리는 중이거나 다른 PC 가 맡았으면 null (나중에 다시).
    /// </summary>
    /// <param name="photoCover">사진 구역·사진 칸용 (경로, 너비, 높이) → JPEG. WPF 쪽에서 준다</param>
    /// <param name="photoFill">사진대지용 (경로, 너비, 높이) → JPEG</param>
    public static PhoneResult? Process(string folder, AppConfig cfg, string baseDir, string machine, DateTime now,
                                       Func<string, int, int, byte[]>? photoCover,
                                       Func<string, int, int, byte[]>? photoFill)
    {
        if (ClaimedByOther(folder, machine, now)) return null;

        var read = PhoneSubmissionReader.Read(folder);
        if (read.Submission == null)
        {
            if (read.MissingPhotos && now - Directory.GetCreationTime(folder) < PhotoWait) return null;
            return Finish(folder, "failed", machine, now, [], read.Errors);
        }
        Claim(folder, machine, now);

        var sub      = read.Submission;
        var messages = new List<string>();
        var docs     = cfg.Documents.Where(d => sub.Docs.Contains(d.Id)).ToList();
        foreach (var id in sub.Docs.Where(id => docs.All(d => d.Id != id)))
            messages.Add($"설정에 없는 문서라 건너뜀: {id}");
        if (docs.Count == 0) return Finish(folder, "failed", machine, now, [], [.. messages, "만들 문서가 없습니다"]);

        var input = new DailyInput { Date = sub.Date, Blocks = PhoneDaily.InitialBlocks(cfg, baseDir, sub.Date, messages) };
        var fields = cfg.Documents.SelectMany(d => d.Fields).Where(f => BlockTypes.IsBlock(f.Type))
                        .GroupBy(BlockTypes.KeyOf).ToDictionary(g => g.Key, g => g.First());
        var photoFields = cfg.Documents.SelectMany(d => d.Fields.Where(f => f.Type == FieldTypes.Photo).Select(f => (d.Id, f.Id)))
                             .ToHashSet();
        PhoneSubmissionReader.ApplyTo(input, sub, fields, photoFields, messages);

        var files = new List<string>();
        int failed = 0;
        foreach (var doc in docs)
        {
            try
            {
                var problems = new List<string>();
                string path = GenericExcelWriter.Generate(input, doc, baseDir, cfg.OutputFolder, photoCover, messages, problems);
                files.Add(Path.GetFileName(path));
                messages.AddRange(problems.Select(p => $"[{doc.Name}] 확인 필요: {p}"));
                InsertSlots(sub, doc, path, baseDir, photoFill, messages);
            }
            catch (Exception ex)
            {
                failed++;
                AppLog.Error($"폰 입력 자동 생성 실패: {doc.Name}", ex);
                messages.Add($"[{doc.Name}] 만들지 못했습니다: {ex.Message}");
            }
        }
        string status = failed == 0 ? "done" : failed == docs.Count ? "failed" : "partial";
        AppLog.Info($"폰 입력 자동 생성 {status}: {Path.GetFileName(folder)} → {string.Join(", ", files)}");
        return Finish(folder, status, machine, now, files, messages);
    }

    private static void InsertSlots(PhoneSubmission sub, DocumentConfig doc, string outputPath, string baseDir,
                                    Func<string, int, int, byte[]>? photoFill, List<string> messages)
    {
        var slots = sub.Slots.Where(s => s.Doc == doc.Id).ToList();
        if (slots.Count == 0) return;
        if (photoFill == null) { messages.Add($"[{doc.Name}] 사진대지는 이 PC 에서 넣을 수 없습니다"); return; }

        string template = Path.Combine(baseDir, doc.TemplateFile);
        foreach (var s in slots)
        {
            try
            {
                var info = ExcelPhotoWriter.GetPhotoSlotInfos(template, s.Sheet).FirstOrDefault(i => i.SlotIndex == s.Index);
                if (info == null) { messages.Add($"[{doc.Name}] 사진대지 {s.Index}번 칸을 찾지 못했습니다"); continue; }
                ExcelPhotoWriter.ReplacePhoto(outputPath, s.Sheet, s.Index, photoFill(s.Path, info.WidthPx, info.HeightPx));
            }
            catch (Exception ex)
            {
                AppLog.Error($"폰 사진대지 넣기 실패: {doc.Name} {s.Index}", ex);
                messages.Add($"[{doc.Name}] 사진대지 {s.Index}번 사진을 넣지 못했습니다: {ex.Message}");
            }
        }
    }

    private static PhoneResult Finish(string folder, string status, string machine, DateTime now,
                                      List<string> files, IEnumerable<string> messages)
    {
        var result = new PhoneResult
        {
            Status = status, At = now.ToString("o", CultureInfo.InvariantCulture), By = machine,
            Files = files, Messages = messages.Distinct().Take(30).ToList(),
        };
        File.WriteAllText(Path.Combine(folder, RESULT_FILE_NAME), JsonSerializer.Serialize(result, PhoneDaily.Json));
        return result;
    }

    private static bool ClaimedByOther(string folder, string machine, DateTime now)
    {
        string path = Path.Combine(folder, CLAIM_FILE_NAME);
        if (!File.Exists(path)) return false;
        try
        {
            var parts = File.ReadAllText(path).Split('|');
            if (parts.Length != 2 || parts[0] == machine) return false;
            return DateTime.TryParse(parts[1], CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var at)
                   && now - at < ClaimValid;
        }
        catch (IOException) { return false; }
    }

    private static void Claim(string folder, string machine, DateTime now) =>
        File.WriteAllText(Path.Combine(folder, CLAIM_FILE_NAME), $"{machine}|{now.ToString("o", CultureInfo.InvariantCulture)}");
}
