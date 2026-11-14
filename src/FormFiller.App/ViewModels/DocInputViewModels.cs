using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Core.Models;
using Microsoft.Win32;

namespace FormFiller.App.ViewModels;

/// <summary>
/// 메인 화면 "문서별 입력" 카드 하나 = 문서 하나.
/// 문서에 매일 입력하는 항목(input / photo / shape_position)이 있을 때만 만든다.
/// 시트가 여러 개(같은 모양)면 시트마다 탭으로 나눠 따로 입력한다.
/// </summary>
public class DocInputCardViewModel : ViewModelBase
{
    public DocumentConfig Config { get; }
    public string Name => Config.Name;
    public ObservableCollection<SheetInputViewModel> Sheets { get; } = [];

    private SheetInputViewModel? _selectedSheet;
    public SheetInputViewModel? SelectedSheet { get => _selectedSheet; set => SetField(ref _selectedSheet, value); }

    public bool HasPhotos => Config.Fields.Any(f => f.Type == FieldTypes.Photo);

    private bool _isVisible = true;
    /// <summary>"오늘 생성할 문서"에서 이 문서를 체크했을 때만 보인다</summary>
    public bool IsVisible { get => _isVisible; set => SetField(ref _isVisible, value); }

    /// <summary>모든 시트의 매일 입력 항목</summary>
    public IEnumerable<(string Sheet, FieldInputViewModel Item)> AllItems =>
        Sheets.SelectMany(s => s.Items.Select(i => (s.Header, i)));

    /// <summary>사진을 가져올 폴더 (비어 있으면 처음 불러올 때 고른다)</summary>
    public string PhotoFolderText => string.IsNullOrEmpty(Config.PhotoFolder) ? "사진 폴더 미지정" : Config.PhotoFolder;

    public ICommand ImportPhotosCommand { get; }
    public ICommand ChangeFolderCommand { get; }

    private readonly Func<DateTime> _getDate;

    public DocInputCardViewModel(DocumentConfig config, Func<DateTime> getDate)
    {
        Config   = config;
        _getDate = getDate;

        foreach (var sheet in config.TargetSheets())
            Sheets.Add(new SheetInputViewModel(config, sheet));
        SelectedSheet = Sheets.FirstOrDefault();

        ImportPhotosCommand = new RelayCommand(_ => ImportPhotos());
        ChangeFolderCommand = new RelayCommand(_ => { if (PickFolder()) ImportPhotos(); });
    }

    /// <summary>입력된 값을 DailyInput.DocValues 에 넣는다.</summary>
    public void FillInput(DailyInput input)
    {
        foreach (var s in Sheets)
            foreach (var f in s.Items)
                f.FillInput(input, Config.Id, s.SheetName);
    }

    // ── 사진 폴더에서 자동으로 불러오기 ───────────────────────────────

    private bool PickFolder()
    {
        var dlg = new OpenFolderDialog { Title = $"{Name} 사진을 올리는 폴더 (예: 구글 드라이브 폴더)" };
        if (dlg.ShowDialog() != true) return false;

        Config.PhotoFolder = dlg.FolderName;
        // 다음에도 쓰도록 config.json 에 저장
        var cfg = AppConfigLoader.Load();
        var doc = cfg.Documents.FirstOrDefault(d => d.Id == Config.Id);
        if (doc != null) { doc.PhotoFolder = dlg.FolderName; AppConfigLoader.Save(cfg); }
        OnPropertyChanged(nameof(PhotoFolderText));
        return true;
    }

    /// <summary>
    /// 사진 폴더에서 "선택한 날짜에 찍은 사진"을 찾아 칸에 넣는다.
    /// - 시트가 여러 개면: 시트 이름이 들어간 하위 폴더의 사진만 그 시트에 쓴다 (예: 폴더 "오전초등학교")
    /// - 칸에 시각(예: 10:00)이 정해져 있으면 촬영 시각이 가장 가까운 사진(±40분 안)을 넣는다
    /// - 시각이 없으면 찍은 순서대로 넣는다
    /// </summary>
    private void ImportPhotos()
    {
        if ((string.IsNullOrEmpty(Config.PhotoFolder) || !Directory.Exists(Config.PhotoFolder)) && !PickFolder())
            return;

        DateTime day = _getDate().Date;
        var exts = new[] { ".jpg", ".jpeg", ".png" };
        var all = Directory.EnumerateFiles(Config.PhotoFolder, "*.*", SearchOption.AllDirectories)
            .Where(f => exts.Contains(Path.GetExtension(f).ToLowerInvariant()))
            .Select(f => (Path: f, Taken: PhotoService.GetTakenTime(f)))
            .Where(p => p.Taken.Date == day)
            .OrderBy(p => p.Taken)
            .ToList();

        int heic = Directory.EnumerateFiles(Config.PhotoFolder, "*.heic", SearchOption.AllDirectories).Count();
        var report = new List<string>();

        foreach (var sheet in Sheets)
        {
            var photoItems = sheet.Items.Where(i => i.IsPhoto).ToList();
            if (photoItems.Count == 0) continue;

            // 시트가 여러 개면 하위 폴더 이름으로 구분
            var candidates = Sheets.Count == 1 ? all
                : all.Where(p => FolderMatches(p.Path, sheet.SheetName)).ToList();

            int placed = 0;
            var used = new HashSet<string>();
            bool hasTimes = photoItems.All(i => i.Time.HasValue);

            if (hasTimes)
            {
                foreach (var item in photoItems)
                {
                    var target = day + item.Time!.Value;
                    var best = candidates.Where(p => !used.Contains(p.Path))
                        .OrderBy(p => Math.Abs((p.Taken - target).TotalMinutes))
                        .FirstOrDefault();
                    if (best.Path != null && Math.Abs((best.Taken - target).TotalMinutes) <= 40)
                    {
                        item.SetPhoto(best.Path); used.Add(best.Path); placed++;
                    }
                }
            }
            else
            {
                foreach (var (item, p) in photoItems.Zip(candidates))
                { item.SetPhoto(p.Path); placed++; }
            }

            report.Add($"{sheet.SheetName}: {placed}/{photoItems.Count}칸 (후보 사진 {candidates.Count}장)");
        }

        string msg = $"{day:yyyy-MM-dd} 에 찍은 사진 {all.Count}장을 찾았습니다.\n\n" + string.Join("\n", report);
        if (Sheets.Count > 1)
            msg += "\n\n※ 시트가 여러 개라서, 시트 이름이 들어간 폴더(예: \"" + Sheets[0].SheetName + "\")의 사진만 그 시트에 넣었습니다.";
        if (heic > 0)
            msg += $"\n\n※ HEIC 사진 {heic}장은 읽을 수 없습니다. 폰 카메라 설정을 '호환성 우선(JPG)'으로 바꿔 주세요.";
        msg += "\n\n잘못 들어간 사진은 사진 칸을 눌러 바꿀 수 있습니다.";
        MessageBox.Show(msg, "사진 불러오기", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private static string Norm(string s) => s.Replace(" ", "").ToLowerInvariant();

    /// <summary>사진 경로의 폴더 이름 중 하나가 시트 이름과 겹치는지 (한쪽이 다른 쪽을 포함)</summary>
    private bool FolderMatches(string path, string sheet)
    {
        string rel = Path.GetRelativePath(Config.PhotoFolder, Path.GetDirectoryName(path) ?? "");
        return rel.Split(Path.DirectorySeparatorChar)
                  .Where(d => d != ".")
                  .Any(d => Norm(d).Contains(Norm(sheet)) || (Norm(d).Length >= 2 && Norm(sheet).Contains(Norm(d))));
    }
}

/// <summary>카드 안의 탭 하나 = 시트 하나</summary>
public class SheetInputViewModel
{
    public string SheetName { get; }
    public string Header => string.IsNullOrEmpty(SheetName) ? "입력" : SheetName;
    public ObservableCollection<FieldInputViewModel> Items { get; } = [];

    public SheetInputViewModel(DocumentConfig doc, string sheet)
    {
        SheetName = sheet;
        foreach (var f in doc.Fields.Where(f => FieldTypes.IsDailyInput(f.Type)))
            Items.Add(new FieldInputViewModel(f));
    }
}

/// <summary>
/// 매일 입력 항목 하나 (직접 입력 / 사진+측정값 / 위치 선택).
/// 화면에서는 타입에 따라 보이는 부분만 다르다 (IsInput / IsPhoto / IsShape).
/// </summary>
public class FieldInputViewModel : ViewModelBase
{
    public FieldConfig Field { get; }
    public string Label  => Field.Label;
    public string Suffix => Field.Suffix ?? "";

    public bool IsInput => Field.Type == FieldTypes.Input;
    public bool IsPhoto => Field.Type == FieldTypes.Photo;
    public bool IsShape => Field.Type == FieldTypes.ShapePosition;

    /// <summary>사진 칸 옆에 측정값 입력 칸을 보여 줄지</summary>
    public bool HasValue => IsInput || (IsPhoto && !string.IsNullOrEmpty(Field.ValueCell));

    public TimeSpan? Time => TimeSpan.TryParse(Field.Time, out var t) ? t : null;
    public string TimeText => Time is { } t ? $"{t:hh\\:mm}" : "";

    // ── 입력 값 ──────────────────────────────────────────────────────

    private string _value = "";
    public string Value { get => _value; set => SetField(ref _value, value); }

    private string? _photoPath;
    public string? PhotoPath
    {
        get => _photoPath;
        private set
        {
            if (SetField(ref _photoPath, value))
            {
                Thumbnail = value == null ? null : PhotoService.LoadOrientedThumbnail(value);
                OnPropertyChanged(nameof(HasPhoto));
                OnPropertyChanged(nameof(PhotoInfo));
            }
        }
    }

    public bool HasPhoto => _photoPath != null;

    /// <summary>사진 아래에 보여 줄 촬영 시각</summary>
    public string PhotoInfo => _photoPath == null ? "" : $"촬영 {PhotoService.GetTakenTime(_photoPath):HH:mm}";

    private BitmapSource? _thumbnail;
    public BitmapSource? Thumbnail { get => _thumbnail; private set => SetField(ref _thumbnail, value); }

    public List<string> PresetNames { get; }

    private string? _selectedPreset;
    public string? SelectedPreset { get => _selectedPreset; set => SetField(ref _selectedPreset, value); }

    private bool _isReading;
    public bool IsReading { get => _isReading; set { if (SetField(ref _isReading, value)) OnPropertyChanged(nameof(ReadButtonText)); } }
    public string ReadButtonText => _isReading ? "읽는 중…" : "사진에서 읽기";

    // ── 커맨드 ───────────────────────────────────────────────────────

    public ICommand BrowseCommand { get; }
    public ICommand ClearCommand  { get; }
    public ICommand ReadValueCommand { get; }

    public FieldInputViewModel(FieldConfig field)
    {
        Field       = field;
        PresetNames = field.Presets?.Select(p => p.Name).ToList() ?? [];

        BrowseCommand = new RelayCommand(_ =>
        {
            var dlg = new OpenFileDialog { Title = $"{Label} 사진 선택", Filter = "사진 (*.jpg;*.jpeg;*.png)|*.jpg;*.jpeg;*.png" };
            if (dlg.ShowDialog() == true) SetPhoto(dlg.FileName);
        });
        ClearCommand     = new RelayCommand(_ => PhotoPath = null);
        ReadValueCommand = new RelayCommand(async _ => await ReadValueAsync());
    }

    public void SetPhoto(string path) => PhotoPath = path;

    /// <summary>사진 속 측정기 숫자를 Claude 로 읽어 입력 칸에 "제안"으로 채운다.</summary>
    private async Task ReadValueAsync()
    {
        if (_photoPath == null || IsReading) return;
        try
        {
            IsReading = true;
            string key = AppConfigLoader.Load().ClaudeApiKey;
            Value = await MeterReader.ReadAsync(key, _photoPath, Field.Suffix);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "읽기 실패", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        finally { IsReading = false; }
    }

    public void FillInput(DailyInput input, string docId, string sheet)
    {
        if (IsInput && Value.Trim().Length > 0)
            input.DocValues[DailyInput.Key(docId, sheet, Field.Id)] = Value.Trim();

        if (IsPhoto)
        {
            if (_photoPath != null)
                input.DocValues[DailyInput.Key(docId, sheet, Field.Id)] = _photoPath;
            if (Value.Trim().Length > 0)
                input.DocValues[DailyInput.Key(docId, sheet, Field.Id + DailyInput.ValueSuffixKey)] = Value.Trim();
        }

        if (IsShape && !string.IsNullOrEmpty(SelectedPreset))
            input.DocValues[DailyInput.Key(docId, sheet, Field.Id)] = SelectedPreset;
    }
}
