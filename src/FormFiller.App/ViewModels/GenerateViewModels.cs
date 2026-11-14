using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using FormFiller.Core.Config;

namespace FormFiller.App.ViewModels;

/// <summary>
/// [파일 생성] 을 누르면 뜨는 창의 ViewModel.
/// 1단계(확인): 문서마다 빠진 칸·문제를 보여 준다.
/// 2단계(결과): 만든 파일 목록과 [열기] [폴더 열기] 버튼을 보여 준다.
/// 같은 창에서 IsDone 값으로 두 화면을 바꿔 보여 준다.
/// </summary>
public class GenerateViewModel : ViewModelBase
{
    public string Title { get; }
    public ObservableCollection<DocCheckItem> Docs { get; }

    private readonly Func<IReadOnlyList<DocCheckItem>, Task<List<string>>> _run;
    private readonly string _outputDir;

    public GenerateViewModel(DateTime date, IEnumerable<DocCheckItem> docs, string outputDir,
                             Func<IReadOnlyList<DocCheckItem>, Task<List<string>>> run)
    {
        Title      = $"{date:yyyy-MM-dd (ddd)} 파일 생성";
        Docs       = new ObservableCollection<DocCheckItem>(docs);
        _outputDir = outputDir;
        _run       = run;
        _openAfter = AppConfigLoader.Load().OpenAfterGenerate;

        GenerateCommand   = new RelayCommand(_ => Generate(), _ => !IsDone && !IsRunning && GeneratableCount > 0);
        OpenFolderCommand = new RelayCommand(_ => OpenFolder());
        OpenAllCommand    = new RelayCommand(_ => { foreach (var d in Docs) d.Open(); });
    }

    // ── 1단계: 확인 ──────────────────────────────────────────────────

    /// <summary>오류(서식 없음, 파일이 열려 있음)가 없는 문서 수 = 실제로 만들 문서 수</summary>
    public int GeneratableCount => Docs.Count(d => !d.HasError);
    public int WarningCount     => Docs.Sum(d => d.Issues.Count);

    public string CheckSummary => WarningCount == 0
        ? "✔ 확인할 것이 없습니다. 바로 생성하세요."
        : $"확인할 것 {WarningCount}건 · 빨간 항목이 있는 문서는 만들지 않습니다.";

    public string GenerateButtonText => _isRunning ? "⏳ 만드는 중…" : $"📄  {GeneratableCount}개 생성";

    private bool _isRunning;
    /// <summary>만드는 중 (창을 닫거나 다시 누르지 못하게)</summary>
    public bool IsRunning
    {
        get => _isRunning;
        private set
        {
            if (!SetField(ref _isRunning, value)) return;
            OnPropertyChanged(nameof(GenerateButtonText));
            CommandManager.InvalidateRequerySuggested();
        }
    }

    private bool _openAfter;
    /// <summary>생성 후 바로 Excel 로 열기 (다음에도 기억)</summary>
    public bool OpenAfter
    {
        get => _openAfter;
        set
        {
            if (!SetField(ref _openAfter, value)) return;
            var cfg = AppConfigLoader.Load();
            cfg.OpenAfterGenerate = value;
            AppConfigLoader.Save(cfg);
        }
    }

    // ── 2단계: 결과 ──────────────────────────────────────────────────

    private bool _isDone;
    public bool IsDone
    {
        get => _isDone;
        private set { if (SetField(ref _isDone, value)) OnPropertyChanged(nameof(IsChecking)); }
    }
    public bool IsChecking => !_isDone;

    private string _resultSummary = "";
    public string ResultSummary { get => _resultSummary; private set => SetField(ref _resultSummary, value); }

    private bool _hasFailures;
    public bool HasFailures { get => _hasFailures; private set => SetField(ref _hasFailures, value); }

    private string _notesText = "";
    /// <summary>엔진이 남긴 참고 사항 (빈 칸 등)</summary>
    public string NotesText
    {
        get => _notesText;
        private set { if (SetField(ref _notesText, value)) OnPropertyChanged(nameof(HasNotes)); }
    }
    public bool HasNotes => _notesText.Length > 0;

    public ICommand GenerateCommand   { get; }
    public ICommand OpenFolderCommand { get; }
    public ICommand OpenAllCommand    { get; }

    private async void Generate()
    {
        if (IsRunning) return;
        var targets = Docs.Where(d => !d.HasError).ToList();
        List<string> notes;
        IsRunning = true;
        try { notes = await _run(targets); }
        catch (Exception ex)
        {
            AppLog.Error("파일 생성 중 오류", ex);
            notes = [$"생성 중 오류: {ex.Message}"];
            foreach (var t in targets.Where(t => t.State == DocResultState.Pending)) t.SetFailed(ex.Message);
        }
        finally { IsRunning = false; }

        foreach (var d in Docs.Where(d => d.HasError))
            d.SetSkipped();

        int ok   = Docs.Count(d => d.State == DocResultState.Done);
        int fail = Docs.Count - ok;
        ResultSummary = fail == 0 ? $"✔ {ok}개 파일을 만들었습니다." : $"{ok}개 생성 · {fail}개 못 만듦";
        HasFailures   = fail > 0;
        NotesText = notes.Count == 0 ? ""
            : string.Join("\n", notes.Take(20)) + (notes.Count > 20 ? $"\n… 외 {notes.Count - 20}건" : "");
        IsDone = true;

        if (OpenAfter)
            foreach (var d in Docs) d.Open();
    }

    private void OpenFolder()
    {
        var first = Docs.FirstOrDefault(d => d.State == DocResultState.Done);
        if (first != null)
            Process.Start("explorer.exe", $"/select,\"{first.OutputPath}\"");
        else if (Directory.Exists(_outputDir))
            Process.Start("explorer.exe", $"\"{_outputDir}\"");
    }
}

public enum DocResultState { Pending, Done, Failed, Skipped }

/// <summary>확인 항목 하나. IsError 면 그 문서는 만들지 않는다(빨강), 아니면 확인만(주황).</summary>
public record CheckIssue(bool IsError, string Message)
{
    public string Icon => IsError ? "✖" : "!";
}

/// <summary>생성 창의 문서 한 줄 (확인 결과 + 생성 결과)</summary>
public class DocCheckItem : ViewModelBase
{
    public DocumentConfig Config { get; }
    public string Name => Config.Name;
    public string OutputPath { get; }
    public string OutputFileName => Path.GetFileName(OutputPath);
    public List<CheckIssue> Issues { get; } = [];

    public bool HasError  => Issues.Any(i => i.IsError);
    public bool HasIssues => Issues.Count > 0;
    public bool NoIssues  => Issues.Count == 0;

    public ICommand OpenCommand { get; }

    public DocCheckItem(DocumentConfig config, string outputPath)
    {
        Config      = config;
        OutputPath  = outputPath;
        OpenCommand = new RelayCommand(_ => Open(), _ => State == DocResultState.Done);
    }

    public void Error(string msg) => Issues.Add(new CheckIssue(true, msg));
    public void Warn(string msg)  => Issues.Add(new CheckIssue(false, msg));

    // ── 생성 결과 ────────────────────────────────────────────────────

    private DocResultState _state;
    public DocResultState State
    {
        get => _state;
        private set
        {
            if (!SetField(ref _state, value)) return;
            OnPropertyChanged(nameof(IsOk));
            OnPropertyChanged(nameof(IsFailed));
        }
    }
    public bool IsOk     => _state == DocResultState.Done;
    public bool IsFailed => _state is DocResultState.Failed or DocResultState.Skipped;

    private string _resultText = "";
    public string ResultText { get => _resultText; private set => SetField(ref _resultText, value); }

    public void SetDone()                { State = DocResultState.Done;   ResultText = "✔ " + OutputFileName; }
    public void SetFailed(string reason) { State = DocResultState.Failed; ResultText = "✖ 실패: " + reason; }
    public void SetSkipped()
    {
        State      = DocResultState.Skipped;
        ResultText = "✖ 만들지 않음: " + Issues.First(i => i.IsError).Message;
    }

    /// <summary>만든 파일을 연결된 프로그램(Excel, 한글)으로 연다.</summary>
    public void Open()
    {
        if (State != DocResultState.Done || !File.Exists(OutputPath)) return;
        try { Process.Start(new ProcessStartInfo(OutputPath) { UseShellExecute = true }); }
        catch (Exception ex)
        {
            MessageBox.Show($"파일을 열지 못했습니다.\n{OutputPath}\n\n{ex.Message}", "열기 실패",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }
}
