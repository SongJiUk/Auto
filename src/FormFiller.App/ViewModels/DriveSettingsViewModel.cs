using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows.Input;
using FormFiller.Core.Config;
using Microsoft.Win32;

namespace FormFiller.App.ViewModels;

/// <summary>
/// [☁ 구글 드라이브] 창.
/// 위: 드라이브 앱·폴더·프로그램 위치 점검 결과 / 아래: 출력 폴더·사진 폴더 고르기.
/// 폴더를 고르면 바로 config.json 에 저장하고 다시 점검한다.
/// </summary>
public class DriveSettingsViewModel : ViewModelBase
{
    public ObservableCollection<DriveCheckRow> Checks { get; } = [];
    public ObservableCollection<PhotoFolderRow> PhotoFolders { get; } = [];

    private string _outputFolder = "";
    /// <summary>출력 폴더 (실제 경로로 보여 준다)</summary>
    public string OutputFolder { get => _outputFolder; private set => SetField(ref _outputFolder, value); }

    private bool _isChecking;
    public bool IsChecking
    {
        get => _isChecking;
        private set { if (SetField(ref _isChecking, value)) { OnPropertyChanged(nameof(Summary)); CommandManager.InvalidateRequerySuggested(); } }
    }

    private CheckState? _worst;
    /// <summary>맨 위 한 줄 요약</summary>
    public string Summary => _isChecking ? "⏳ 점검 중…" : _worst switch
    {
        CheckState.Ok      => "✔ 모두 정상입니다. 두 PC 가 같은 설정과 파일을 씁니다.",
        CheckState.Warning => "⚠ 확인할 것이 있습니다. 아래 주황색 줄을 보세요.",
        CheckState.Error   => "✖ 문제가 있습니다. 빨간 줄의 해결 방법을 따라 주세요.",
        _                  => "",
    };
    public string SummaryColor => _worst switch
    {
        CheckState.Ok      => "#276749",
        CheckState.Warning => "#C05621",
        CheckState.Error   => "#C53030",
        _                  => "#4A5568",
    };

    public ICommand RecheckCommand       { get; }
    public ICommand OpenDownloadCommand  { get; }
    public ICommand BrowseOutputCommand  { get; }
    public ICommand OpenOutputCommand    { get; }

    public DriveSettingsViewModel()
    {
        RecheckCommand      = new RelayCommand(_ => _ = RunChecksAsync(), _ => !IsChecking);
        OpenDownloadCommand = new RelayCommand(_ => OpenShell(GoogleDrive.DownloadUrl));
        BrowseOutputCommand = new RelayCommand(_ => BrowseOutput());
        OpenOutputCommand   = new RelayCommand(_ => { if (Directory.Exists(OutputFolder)) OpenShell(OutputFolder); });
        LoadFolders();
        _ = RunChecksAsync();
    }

    /// <summary>설정에서 폴더 목록을 다시 읽어 화면에 채운다.</summary>
    private void LoadFolders()
    {
        var cfg = AppConfigLoader.Load();
        OutputFolder = DriveChecker.ResolveOutputFolder(cfg, AppDomain.CurrentDomain.BaseDirectory);

        PhotoFolders.Clear();
        foreach (var doc in cfg.Documents.Where(d => d.Fields.Any(f => f.Type == FieldTypes.Photo)))
            PhotoFolders.Add(new PhotoFolderRow(doc.Id, doc.Name, doc.PhotoFolder, OnPhotoFolderPicked));
    }

    /// <summary>
    /// 점검은 드라이브 문자를 훑고 쓰기 시험을 해서 조금 걸릴 수 있다 → 화면 스레드 밖에서.
    /// (await: 뒤에서 일이 끝날 때까지 화면을 멈추지 않고 기다렸다가, 끝나면 이어서 화면 스레드에서 실행한다)
    /// </summary>
    public async Task RunChecksAsync()
    {
        if (IsChecking) return;
        IsChecking = true;
        try
        {
            var cfg     = AppConfigLoader.Load();
            string bd   = AppDomain.CurrentDomain.BaseDirectory;
            var items   = await Task.Run(() => DriveChecker.Run(cfg, bd));

            Checks.Clear();
            foreach (var i in items) Checks.Add(new DriveCheckRow(i));
            _worst = items.Count == 0 ? CheckState.Ok : items.Max(i => i.State);
        }
        catch (Exception ex)
        {
            AppLog.Error("구글 드라이브 점검 실패", ex);
            Checks.Clear();
            Checks.Add(new DriveCheckRow(new DriveCheckItem("점검", CheckState.Error, ex.Message)));
            _worst = CheckState.Error;
        }
        finally
        {
            IsChecking = false;
            OnPropertyChanged(nameof(SummaryColor));
        }
    }

    private void BrowseOutput()
    {
        var dlg = new OpenFolderDialog { Title = "만든 파일을 저장할 폴더 (구글 드라이브 안 폴더 권장)" };
        if (Directory.Exists(OutputFolder)) dlg.InitialDirectory = OutputFolder;
        if (dlg.ShowDialog() != true) return;

        var cfg = AppConfigLoader.Load();
        cfg.OutputFolder = AppConfigLoader.ToStoredFolderPath(dlg.FolderName);
        AppConfigLoader.Save(cfg);
        LoadFolders();
        _ = RunChecksAsync();
    }

    private void OnPhotoFolderPicked(string docId, string folder)
    {
        var cfg = AppConfigLoader.Load();
        var doc = cfg.Documents.FirstOrDefault(d => d.Id == docId);
        if (doc == null) return;
        doc.PhotoFolder = folder;
        AppConfigLoader.Save(cfg);
        _ = RunChecksAsync();
    }

    /// <summary>주소나 폴더를 윈도우 기본 프로그램(브라우저·탐색기)으로 연다.</summary>
    private static void OpenShell(string target)
    {
        try { Process.Start(new ProcessStartInfo(target) { UseShellExecute = true }); }
        catch (Exception ex) { AppLog.Error($"열기 실패: {target}", ex); }
    }
}

/// <summary>점검 결과 한 줄 (아이콘·색은 화면 표시용)</summary>
public class DriveCheckRow
{
    public DriveCheckItem Item { get; }
    public DriveCheckRow(DriveCheckItem item) => Item = item;

    public string Title         => Item.Title;
    public string Detail        => Item.Detail;
    public string Fix           => Item.Fix ?? "";
    public bool   HasFix        => !string.IsNullOrEmpty(Item.Fix) && Item.State != CheckState.Ok;
    public bool   OfferDownload => Item.OfferDownload && Item.State != CheckState.Ok;

    public string Icon => Item.State switch { CheckState.Ok => "✔", CheckState.Warning => "⚠", _ => "✖" };
    public string Color => Item.State switch { CheckState.Ok => "#38A169", CheckState.Warning => "#DD6B20", _ => "#E53E3E" };
    public string Background => Item.State switch { CheckState.Ok => "#F0FFF4", CheckState.Warning => "#FFFAF0", _ => "#FFF5F5" };
}

/// <summary>문서 하나의 사진 폴더 줄</summary>
public class PhotoFolderRow : ViewModelBase
{
    private readonly string _docId;
    private readonly Action<string, string> _onPicked;

    public string DocName { get; }

    private string _folder;
    public string Folder
    {
        get => _folder;
        private set { if (SetField(ref _folder, value)) OnPropertyChanged(nameof(FolderText)); }
    }
    public string FolderText => string.IsNullOrEmpty(_folder) ? "(아직 정하지 않음)" : _folder;

    public ICommand BrowseCommand { get; }

    public PhotoFolderRow(string docId, string docName, string folder, Action<string, string> onPicked)
    {
        _docId    = docId;
        DocName   = docName;
        _folder   = folder;
        _onPicked = onPicked;
        BrowseCommand = new RelayCommand(_ => Browse());
    }

    private void Browse()
    {
        var dlg = new OpenFolderDialog { Title = $"{DocName} 사진을 올리는 폴더 (구글 드라이브 안 폴더)" };
        if (Directory.Exists(_folder)) dlg.InitialDirectory = _folder;
        if (dlg.ShowDialog() != true) return;
        Folder = dlg.FolderName;
        _onPicked(_docId, dlg.FolderName);
    }
}
