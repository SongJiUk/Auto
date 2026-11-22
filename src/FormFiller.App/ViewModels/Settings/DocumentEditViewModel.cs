using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using Microsoft.Win32;

namespace FormFiller.App.ViewModels.Settings;

/// <summary>
/// 문서 관리 화면에서 한 문서(DocumentConfig)를 보여 주고, 이름·서식 파일·출력 이름·주기만 고치는 ViewModel.
/// 칸·규칙·사진은 [✏ 칸 편집](DocumentSetupWindow)에서만 바꾼다 (2026-11-22 화면 정리).
/// 이 화면에서 다루지 않는 값(시트·사진·수장재 등)은 ToConfig() 때 원본 그대로 보존한다.
/// </summary>
public class DocumentEditViewModel : ViewModelBase
{
    public string Id { get; }

    private string _name;
    public string Name { get => _name; set { if (SetField(ref _name, value)) OnPropertyChanged(nameof(ListTitle)); } }

    private string _templateFile;
    public string TemplateFile
    {
        get => _templateFile;
        set
        {
            if (!SetField(ref _templateFile, value)) return;
            OnPropertyChanged(nameof(TemplateMissing));
            OnPropertyChanged(nameof(ListSubtitle));
        }
    }

    /// <summary>서식 파일이 없으면 true (경고 표시용)</summary>
    public bool TemplateMissing
    {
        get
        {
            string path = AppConfigLoader.ResolvePath(_templateFile);
            return string.IsNullOrWhiteSpace(path) || !File.Exists(path);
        }
    }

    private string _outputPattern;
    public string OutputPattern { get => _outputPattern; set => SetField(ref _outputPattern, value); }

    // ── 생성 주기 ────────────────────────────────────────────────────

    public IReadOnlyList<FrequencyOption> FrequencyOptions { get; } = FrequencyOption.All;

    private FrequencyOption _selectedFrequency;
    public FrequencyOption SelectedFrequency
    {
        get => _selectedFrequency;
        set
        {
            if (!SetField(ref _selectedFrequency, value)) return;
            OnPropertyChanged(nameof(IsCustomFrequency));
            OnPropertyChanged(nameof(IsWeeklyFrequency));
            OnPropertyChanged(nameof(ListSubtitle));
        }
    }

    public bool IsCustomFrequency => _selectedFrequency?.Value == "custom";

    private int _frequencyDays;
    public int FrequencyDays { get => _frequencyDays; set => SetField(ref _frequencyDays, value); }

    public bool IsWeeklyFrequency => _selectedFrequency?.Value == "weekly";
    public IReadOnlyList<WeekdayOption> WeekdayOptions { get; } = WeekdayOption.All;

    private WeekdayOption _selectedWeekday;
    public WeekdayOption SelectedWeekday { get => _selectedWeekday; set => SetField(ref _selectedWeekday, value); }

    private bool _useWeeklyDayDate;
    /// <summary>true: 그 주 정한 요일 날짜로 파일을 만든다 / false: 만드는 날 날짜</summary>
    public bool UseWeeklyDayDate { get => _useWeeklyDayDate; set => SetField(ref _useWeeklyDayDate, value); }

    // ── 항목 (읽기 전용) ─────────────────────────────────────────────

    public ObservableCollection<FieldEditRow> Fields { get; } = [];

    public bool HasLegacyFields => Fields.Any(f => f.IsLegacy);

    // ── 왼쪽 목록 표시 ───────────────────────────────────────────────

    public string ListTitle => string.IsNullOrWhiteSpace(Name) ? "(이름 없음)" : Name;

    public string ListSubtitle =>
        $"{_selectedFrequency?.Label ?? "매일"} · 항목 {Fields.Count}개" + (TemplateMissing ? " · ⚠ 서식 없음" : "");

    // ── 커맨드 ──────────────────────────────────────────────────────

    public ICommand BrowseFileCommand   { get; }
    public ICommand TestGenerateCommand { get; }

    private readonly DocumentConfig _original;

    public DocumentEditViewModel(DocumentConfig doc)
    {
        _original         = doc;
        Id                = doc.Id;
        _name             = doc.Name;
        _templateFile     = doc.TemplateFile;
        _outputPattern    = doc.OutputPattern;
        _frequencyDays    = doc.FrequencyDays;
        _selectedWeekday  = WeekdayOption.Of(doc.WeeklyDay);
        _useWeeklyDayDate = doc.UseWeeklyDayDate;
        _selectedFrequency = FrequencyOption.All.FirstOrDefault(f => f.Value == doc.Frequency)
                             ?? FrequencyOption.All[0];

        foreach (var f in doc.Fields)
            Fields.Add(new FieldEditRow(f));

        BrowseFileCommand   = new RelayCommand(_ => BrowseFile());
        TestGenerateCommand = new RelayCommand(_ => TestGenerate());
    }

    private void BrowseFile()
    {
        var dlg = new OpenFileDialog
        {
            Title  = "서식 파일 선택",
            Filter = "서식 파일 (*.xlsx;*.hwpx)|*.xlsx;*.hwpx",
        };
        if (dlg.ShowDialog() == true)
            TemplateFile = dlg.FileName;
    }

    /// <summary>[🧪 시험 생성]: 지금 화면 값으로 임시 폴더에 오늘 파일을 만들어 본다 (저장 안 함).</summary>
    private void TestGenerate()
    {
        string template = AppConfigLoader.ResolvePath(_templateFile);
        if (TemplateMissing)
        {
            MessageBox.Show("서식 파일을 찾을 수 없습니다. [찾아보기]로 서식을 다시 골라 주세요.", "시험 생성",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (!template.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show("시험 생성은 엑셀 서식만 됩니다.", "시험 생성", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            var notes = new List<string>();
            string path = TrialGenerator.Generate(ToConfig(), template, TrialGenerator.NewTempFolder(), DateTime.Today, notes);

            string msg = $"시험 파일을 만들었습니다 (임시 폴더, 실제 출력 폴더는 그대로).\n{path}\n\n" +
                         "입력 칸은 서식에 적힌 글자를 그대로 넣었습니다. 칸 위치·날짜 형식·식 결과를 확인하세요.";
            if (notes.Count > 0) msg += "\n\n참고:\n· " + string.Join("\n· ", notes.Take(8));
            if (MessageBox.Show(msg + "\n\n지금 열까요?", "시험 생성", MessageBoxButton.YesNo,
                                MessageBoxImage.Information) == MessageBoxResult.Yes)
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            AppLog.Error("시험 생성 실패", ex);
            MessageBox.Show($"시험 생성 중 오류가 났습니다.\n{ex.Message}", "시험 생성",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    /// <summary>현재 편집 상태를 DocumentConfig 로. 이 화면에서 다루지 않는 값은 원본 그대로.</summary>
    public DocumentConfig ToConfig()
    {
        var doc = System.Text.Json.JsonSerializer.Deserialize<DocumentConfig>(
                      System.Text.Json.JsonSerializer.Serialize(_original))!;
        doc.Name             = Name;
        doc.TemplateFile     = AppConfigLoader.ToStoredTemplatePath(TemplateFile);
        doc.OutputPattern    = OutputPattern;
        doc.Frequency        = _selectedFrequency?.Value ?? "daily";
        if (IsCustomFrequency) doc.FrequencyDays = FrequencyDays;   // 그 외 주기에서는 쓰지 않으므로 원래 값 유지
        doc.WeeklyDay        = SelectedWeekday.Value;
        doc.UseWeeklyDayDate = IsWeeklyFrequency && UseWeeklyDayDate;
        doc.Fields           = Fields.Select(f => f.ToFieldConfig()).ToList();
        return doc;
    }
}
