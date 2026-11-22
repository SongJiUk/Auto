using System.IO;
using System.Windows.Threading;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Core.Phone;

namespace FormFiller.App;

/// <summary>
/// 켜 둔 PC 가 폰 입력을 1분마다 확인해서 문서를 자동으로 만든다 (2026-11-23 사용자 결정).
/// - 폰 화면 양식(phone_form.json)을 드라이브 inbox 에 써 둔다 (날짜가 바뀌거나 10분마다, 만든 뒤에도)
/// - 새 폰 입력이 있으면 PC 설정대로 만들고 result.json 을 남긴다 (폰 화면에서 결과를 본다)
/// 드라이브 밖 복사본(시험용 등)에서는 돌지 않는다 — 실제 출력 폴더를 건드리지 않게.
/// </summary>
public sealed class PhoneWatcher
{
    private static readonly TimeSpan Interval    = TimeSpan.FromMinutes(1);
    private static readonly TimeSpan FormRefresh = TimeSpan.FromMinutes(10);

    private readonly DispatcherTimer _timer = new() { Interval = Interval };
    private readonly Action<string> _report;
    private bool _busy;
    private DateTime _formWrittenAt = DateTime.MinValue;
    private DateTime _formDate = DateTime.MinValue;

    /// <param name="report">만든 결과를 화면에 한 줄로 알리는 곳 (화면 스레드에서 부른다)</param>
    public PhoneWatcher(Action<string> report)
    {
        _report = report;
        _timer.Tick += async (_, _) => await TickAsync();
    }

    public void Start()
    {
        _timer.Start();
        _ = TickAsync();
    }

    /// <summary>설정이 바뀌었을 때 양식을 곧바로 다시 쓰게 한다</summary>
    public void InvalidateForm() => _formWrittenAt = DateTime.MinValue;

    private async Task TickAsync()
    {
        if (_busy) return;
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        if (!GoogleDrive.IsInsideDrive(baseDir)) return;

        _busy = true;
        try
        {
            var lines = await BackgroundWork.RunSta(() => Work(baseDir));
            foreach (var line in lines) _report(line);
        }
        catch (Exception ex)
        {
            AppLog.Error("폰 입력 확인 실패", ex);
        }
        finally { _busy = false; }
    }

    /// <summary>뒤 스레드: 대기 중인 폰 입력 처리 + 양식 갱신. 화면에 알릴 줄을 돌려준다.</summary>
    private List<string> Work(string baseDir)
    {
        var cfg   = AppConfigLoader.Load();
        string inbox = PhoneInbox.ResolveInboxFolder(cfg.PhoneInboxFolder, baseDir);
        var now   = DateTime.Now;
        var lines = new List<string>();

        bool generated = false;
        foreach (var folder in PhoneAutoGenerator.FindPending(inbox))
        {
            var result = PhoneAutoGenerator.Process(folder, cfg, baseDir, Environment.MachineName, now,
                                                    PhotoService.ResizeToCover, PhotoService.ResizeToFill);
            if (result == null) continue;   // 사진 동기화 대기 / 다른 PC 가 처리 중
            generated |= result.Files.Count > 0;
            lines.Add(result.Status switch
            {
                "done"    => $"📱 폰 입력으로 만듦 {now:HH:mm}: {string.Join(", ", result.Files)}",
                "partial" => $"📱 폰 입력 일부만 만듦 {now:HH:mm} — 로그 확인",
                _         => $"📱 폰 입력을 만들지 못함 {now:HH:mm}: {result.Messages.FirstOrDefault()}",
            });
        }

        if (generated || now.Date != _formDate || now - _formWrittenAt > FormRefresh)
        {
            WriteForm(cfg, baseDir, inbox, now);
            _formDate      = now.Date;
            _formWrittenAt = now;
        }
        return lines;
    }

    private static void WriteForm(AppConfig cfg, string baseDir, string inbox, DateTime now)
    {
        DateTime today = now.Date;
        var initial = PhoneDaily.InitialBlocks(cfg, baseDir, today);
        string output = Path.Combine(baseDir, cfg.OutputFolder);
        var form = PhoneDaily.BuildForm(cfg, today, initial,
            doc => DocumentSchedule.IsDue(doc, today, DocumentSchedule.LastGenerated(doc, output, notAfter: today)),
            doc => SlotIndexes(doc, baseDir),
            Environment.MachineName, now);
        PhoneDaily.WriteForm(inbox, form);
    }

    private static IReadOnlyList<int> SlotIndexes(DocumentConfig doc, string baseDir)
    {
        try
        {
            string template = Path.Combine(baseDir, doc.TemplateFile);
            return File.Exists(template)
                ? ExcelPhotoWriter.GetPhotoSlotInfos(template, doc.PhotoSheet).Select(s => s.SlotIndex).ToList()
                : [];
        }
        catch (Exception ex)
        {
            AppLog.Error($"사진대지 칸 읽기 실패: {doc.Name}", ex);
            return [];
        }
    }
}
