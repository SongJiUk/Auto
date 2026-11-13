using System.Collections.ObjectModel;
using System.IO;
using System.Windows.Input;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using Microsoft.Win32;

namespace FormFiller.App.ViewModels.Settings;

/// <summary>
/// 설정 화면에서 한 문서(DocumentConfig)를 편집하는 ViewModel.
/// 생성자에서 DocumentConfig를 받아 프로퍼티에 풀고,
/// ToConfig()로 다시 DocumentConfig로 변환해 저장에 사용한다.
/// </summary>
public class DocumentEditViewModel : ViewModelBase
{
    // ── 식별자 (변경 불가) ────────────────────────────────────────────
    public string Id { get; }

    // ── 기본 정보 ────────────────────────────────────────────────────

    private string _name;
    public string Name { get => _name; set => SetField(ref _name, value); }

    private string _templateFile;

    /// <summary>
    /// 서식 파일 경로. 바뀌면 시트 목록을 자동으로 다시 읽는다.
    /// </summary>
    public string TemplateFile
    {
        get => _templateFile;
        set { if (SetField(ref _templateFile, value)) ReloadSheets(); }
    }

    private string _outputPattern;
    public string OutputPattern { get => _outputPattern; set => SetField(ref _outputPattern, value); }

    // ── 시트 ────────────────────────────────────────────────────────

    public ObservableCollection<string> SheetNames { get; } = [];

    private string _selectedSheet = "";
    public string SelectedSheet { get => _selectedSheet; set => SetField(ref _selectedSheet, value); }

    // ── 생성 주기 ────────────────────────────────────────────────────

    public IReadOnlyList<FrequencyOption> FrequencyOptions { get; } = FrequencyOption.All;

    private FrequencyOption _selectedFrequency;
    public FrequencyOption SelectedFrequency
    {
        get => _selectedFrequency;
        set { if (SetField(ref _selectedFrequency, value)) OnPropertyChanged(nameof(IsCustomFrequency)); }
    }

    public bool IsCustomFrequency => _selectedFrequency?.Value == "custom";

    private int _frequencyDays;
    public int FrequencyDays { get => _frequencyDays; set => SetField(ref _frequencyDays, value); }

    // ── 수장재팀 포함 여부 ────────────────────────────────────────────

    private bool _includesSujangjae;
    public bool IncludesSujangjae { get => _includesSujangjae; set => SetField(ref _includesSujangjae, value); }

    // ── 사진 시트 ────────────────────────────────────────────────────

    private const string NoPhotoOption = "없음 (사진 기능 미사용)";

    /// <summary>"없음…" + SheetNames 로 구성된 사진 시트 선택 목록</summary>
    public ObservableCollection<string> PhotoSheetOptions { get; } = [NoPhotoOption];

    private string _selectedPhotoSheet = NoPhotoOption;

    public string SelectedPhotoSheet
    {
        get => _selectedPhotoSheet;
        set
        {
            if (SetField(ref _selectedPhotoSheet, value))
            {
                OnPropertyChanged(nameof(PhotoSheet));
                ReloadPhotoSlotRows();
            }
        }
    }

    // ── 사진 슬롯 이름 목록 ──────────────────────────────────────────

    /// <summary>
    /// 선택된 사진 시트에서 감지된 슬롯 목록.
    /// 각 행에서 사용자가 슬롯 이름을 편집할 수 있다.
    /// </summary>
    public ObservableCollection<PhotoSlotLabelRow> PhotoSlotRows { get; } = [];

    /// <summary>저장용 실제 값. "없음…" 선택이면 빈 문자열.</summary>
    public string PhotoSheet =>
        _selectedPhotoSheet == NoPhotoOption ? "" : _selectedPhotoSheet ?? "";

    // ── 필드 목록 ────────────────────────────────────────────────────

    public ObservableCollection<FieldEditRow> Fields { get; } = [];

    // ── 커맨드 ──────────────────────────────────────────────────────

    public ICommand BrowseFileCommand     { get; }
    public ICommand DeleteDocumentCommand { get; }

    // ── 내부 ────────────────────────────────────────────────────────

    /// <summary>config에서 불러온 초기 슬롯 이름 (index → label). 슬롯 로드 시 복원에 사용.</summary>
    private readonly Dictionary<int, string> _initialSlotLabels;

    // ── 생성자 ──────────────────────────────────────────────────────

    /// <summary>불러온 원본 (이 화면에서 다루지 않는 값 — 여러 시트, 사진 폴더 등 — 을 저장 때 그대로 보존)</summary>
    private readonly DocumentConfig _original;

    public DocumentEditViewModel(DocumentConfig doc, Action<DocumentEditViewModel> onDelete)
    {
        _original          = doc;
        Id                 = doc.Id;
        _name              = doc.Name;
        _templateFile      = doc.TemplateFile;
        _outputPattern     = doc.OutputPattern;
        _selectedSheet     = doc.SheetName;
        _frequencyDays     = doc.FrequencyDays;
        _includesSujangjae = doc.IncludesSujangjae;

        // 저장된 photoSheet 값을 선택 상태로 복원 (ReloadSheets 후에 처리)
        _selectedPhotoSheet = string.IsNullOrEmpty(doc.PhotoSheet) ? NoPhotoOption : doc.PhotoSheet;

        _selectedFrequency = FrequencyOption.All.FirstOrDefault(f => f.Value == doc.Frequency)
                             ?? FrequencyOption.All[0];

        _initialSlotLabels = doc.PhotoSlots.ToDictionary(s => s.Index, s => s.Label);

        foreach (var f in doc.Fields)
            Fields.Add(new FieldEditRow(f));

        ReloadSheets();
        // ReloadSheets 가 PhotoSheetOptions 를 채운 뒤 슬롯도 로드한다
        ReloadPhotoSlotRows();

        BrowseFileCommand     = new RelayCommand(_ => BrowseFile());
        DeleteDocumentCommand = new RelayCommand(_ => onDelete(this));
    }

    // ── 메서드 ──────────────────────────────────────────────────────

    private void BrowseFile()
    {
        var dlg = new OpenFileDialog
        {
            Title  = "서식 파일 선택",
            Filter = "Excel 파일 (*.xlsx)|*.xlsx",
        };
        if (dlg.ShowDialog() == true)
            TemplateFile = dlg.FileName;
    }

    private void ReloadSheets()
    {
        SheetNames.Clear();
        if (!string.IsNullOrWhiteSpace(_templateFile) && File.Exists(_templateFile))
        {
            foreach (var s in ExcelAnalyzer.GetSheetNames(_templateFile))
                SheetNames.Add(s);

            if (!SheetNames.Contains(_selectedSheet) && SheetNames.Count > 0)
                SelectedSheet = SheetNames[0];
        }

        // 사진 시트 선택 목록 갱신
        PhotoSheetOptions.Clear();
        PhotoSheetOptions.Add(NoPhotoOption);
        foreach (var s in SheetNames)
            PhotoSheetOptions.Add(s);

        // 이전 선택값이 목록에 있으면 유지, 없으면 "없음"으로
        if (!PhotoSheetOptions.Contains(_selectedPhotoSheet))
            SelectedPhotoSheet = NoPhotoOption;
    }

    /// <summary>
    /// 선택된 사진 시트에서 이미지 슬롯을 감지해 PhotoSlotRows를 갱신한다.
    /// 이전 행에 이름이 있으면 최대한 보존한다.
    /// </summary>
    private void ReloadPhotoSlotRows()
    {
        // 현재 편집 중인 이름을 보존 (시트가 바뀌면 버리고 초기값으로 돌아감)
        var currentLabels = PhotoSlotRows.ToDictionary(r => r.SlotIndex, r => r.Label);
        PhotoSlotRows.Clear();

        if (string.IsNullOrEmpty(PhotoSheet)
            || string.IsNullOrWhiteSpace(_templateFile)
            || !File.Exists(_templateFile)) return;

        try
        {
            var slotInfos = ExcelPhotoWriter.GetPhotoSlotInfos(_templateFile, PhotoSheet);
            foreach (var slot in slotInfos)
            {
                // 우선순위: 현재 편집값 > config 저장값 > 빈 문자열
                string label = currentLabels.TryGetValue(slot.SlotIndex, out var cur) ? cur
                             : _initialSlotLabels.TryGetValue(slot.SlotIndex, out var init) ? init
                             : "";
                byte[]? imgBytes = ExcelPhotoWriter.GetSlotImageBytes(_templateFile, PhotoSheet, slot.SlotIndex);
                PhotoSlotRows.Add(new PhotoSlotLabelRow(slot.SlotIndex, slot.WidthPx, slot.HeightPx, label, imgBytes));
            }
        }
        catch { /* 파일 접근 실패 시 목록 비움 */ }
    }

    /// <summary>현재 편집 상태를 DocumentConfig로 변환한다.</summary>
    public DocumentConfig ToConfig() => new()
    {
        Id                = Id,
        Name              = Name,
        TemplateFile      = TemplateFile,
        OutputPattern     = OutputPattern,
        SheetName         = SelectedSheet,
        Frequency         = _selectedFrequency?.Value ?? "daily",
        FrequencyDays     = IsCustomFrequency ? FrequencyDays : 0,
        IncludesSujangjae = IncludesSujangjae,
        PhotoSheet        = PhotoSheet,
        PhotoSlots        = PhotoSlotRows.Select(r => r.ToConfig()).ToList(),
        Fields            = Fields.Select(f => f.ToFieldConfig()).ToList(),
        // 이 화면에서 편집하지 않는 값은 원본 그대로
        SheetNames        = _original.SheetNames.Contains(SelectedSheet) ? _original.SheetNames : [],
        PhotoFolder       = _original.PhotoFolder,
        SaveMode          = _original.SaveMode,
    };
}
