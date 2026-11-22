using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using Microsoft.Win32;

namespace FormFiller.App.ViewModels.Settings;

/// <summary>
/// 문서 설정 창(DocumentSetupWindow)의 ViewModel. 새 문서 추가와 기존 문서 편집을 함께 맡는다.
///
/// [사용 흐름]
/// 1. 서식 파일(.xlsx 또는 .hwpx)을 고른다 → 왼쪽에 Excel 이 열린다.
/// 2. "+ 항목 추가"에서 이름과 종류를 고르면, 바로 "Excel 에서 클릭하세요" 상태가 된다.
/// 3. Excel 에서 칸·사진·그림을 클릭(또는 드래그 후 확인)하면 지정이 끝난다.
/// 4. 같은 모양 시트가 여러 개면 "적용할 시트"에서 모두 고른다 → 한 번 설정으로 전부 채운다.
///
/// Excel 조작은 IExcelBridge(View 가 구현)를 통해서만 한다.
/// </summary>
public class DocumentSetupViewModel : ViewModelBase, ISetupHost
{
    /// <summary>View 가 연결해 주는 Excel 창구. hwpx 이거나 Excel 이 없으면 null.</summary>
    public IExcelBridge? Bridge { get; set; }

    /// <summary>편집 중인 기존 문서 (새 문서면 null)</summary>
    private DocumentConfig? _editing;

    public string WindowTitle => _editing == null ? "새 문서 설정" : $"문서 편집 — {_editing.Name}";

    // ════════════════════════════════════════════════════════════════
    // ── 파일 / 시트 ───────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    private string _filePath = "";

    /// <summary>선택된 서식 파일 경로. 바뀌면 시트 목록을 다시 읽는다.</summary>
    public string FilePath
    {
        get => _filePath;
        set
        {
            if (SetField(ref _filePath, value))
            {
                OnPropertyChanged(nameof(IsHwpx));
                OnPropertyChanged(nameof(IsExcel));
                OnPropertyChanged(nameof(HasFile));
                foreach (var i in Items) i.RefreshMode();
                ReloadSheets();
            }
        }
    }

    public bool HasFile => File.Exists(_filePath);
    public bool IsHwpx  => HwpxWriter.IsHwpx(_filePath);
    public bool IsExcel => !IsHwpx;

    /// <summary>파일의 시트 이름 목록 (미리보기 시트 선택용)</summary>
    public ObservableCollection<string> SheetNames { get; } = [];

    /// <summary>"적용할 시트" 칩 목록 (같은 모양 시트 여러 개 선택)</summary>
    public ObservableCollection<SheetChoice> SheetChoices { get; } = [];

    private string _selectedSheet = "";

    /// <summary>왼쪽 Excel 에 보이는 시트. 바뀌면 View 가 Excel 시트를 전환한다.</summary>
    public string SelectedSheet
    {
        get => _selectedSheet;
        set { if (SetField(ref _selectedSheet, value)) RefreshHighlights(); }   // 시트를 바꾸면 그 시트의 칸을 칠한다
    }

    // ════════════════════════════════════════════════════════════════
    // ── 문서 기본 정보 ────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    private string _documentName = "";
    public string DocumentName { get => _documentName; set => SetField(ref _documentName, value); }

    private string _outputPattern = "";

    /// <summary>출력 파일명 패턴. {yy}=연도 2자리, {MM}=월, {dd}=일</summary>
    public string OutputPattern { get => _outputPattern; set => SetField(ref _outputPattern, value); }

    public IReadOnlyList<FrequencyOption> FrequencyOptions { get; } = FrequencyOption.All;

    private FrequencyOption _selectedFrequency = FrequencyOption.All[0];
    public FrequencyOption SelectedFrequency
    {
        get => _selectedFrequency;
        set
        {
            if (!SetField(ref _selectedFrequency, value)) return;
            OnPropertyChanged(nameof(IsCustomFrequency));
            OnPropertyChanged(nameof(IsWeeklyFrequency));
        }
    }

    public bool IsCustomFrequency => _selectedFrequency?.Value == "custom";

    private int _frequencyDays = 7;
    public int FrequencyDays { get => _frequencyDays; set => SetField(ref _frequencyDays, value); }

    // ── 매주 주기: 요일 + 파일 날짜 ──────────────────────────────────
    public bool IsWeeklyFrequency => _selectedFrequency?.Value == "weekly";
    public IReadOnlyList<WeekdayOption> WeekdayOptions { get; } = WeekdayOption.All;

    private WeekdayOption _selectedWeekday = WeekdayOption.Of(DayOfWeek.Monday);
    public WeekdayOption SelectedWeekday { get => _selectedWeekday; set => SetField(ref _selectedWeekday, value); }

    private bool _useWeeklyDayDate;
    /// <summary>true: 그 주 정한 요일 날짜로 파일을 만든다 / false: 만드는 날 날짜</summary>
    public bool UseWeeklyDayDate { get => _useWeeklyDayDate; set => SetField(ref _useWeeklyDayDate, value); }

    private bool _includesSujangjae;
    public bool IncludesSujangjae { get => _includesSujangjae; set { if (SetField(ref _includesSujangjae, value)) OnPropertyChanged(nameof(ShowSujangjae)); } }

    /// <summary>수장재 체크는 예전 방식 공사일보용 → 새 문서에서는 숨긴다 (이미 켠 문서·예전 항목이 있을 때만)</summary>
    public bool ShowSujangjae => _includesSujangjae || HasLegacyItems;

    /// <summary>예전 사진대지 슬롯은 이미 쓰는 문서에서만 보인다 (새 문서는 사진 자리 항목으로)</summary>
    public bool ShowLegacyPhotoSlots => IsExcel && PhotoSheet.Length > 0;

    // ── 사진대지 슬롯 (예전 방식: 시트의 그림을 순서대로 교체) ──────────

    private const string NoPhotoOption = "없음";
    public ObservableCollection<string> PhotoSheetOptions { get; } = [NoPhotoOption];

    private string _selectedPhotoSheet = NoPhotoOption;
    public string SelectedPhotoSheet
    {
        get => _selectedPhotoSheet;
        set { if (SetField(ref _selectedPhotoSheet, value)) { ReloadPhotoSlotRows(); OnPropertyChanged(nameof(ShowLegacyPhotoSlots)); } }
    }

    public string PhotoSheet => _selectedPhotoSheet == NoPhotoOption ? "" : _selectedPhotoSheet ?? "";
    public ObservableCollection<PhotoSlotLabelRow> PhotoSlotRows { get; } = [];
    private Dictionary<int, string> _initialSlotLabels = [];

    // ════════════════════════════════════════════════════════════════
    // ── 항목 목록 ─────────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    public ObservableCollection<SetupItemViewModel> Items { get; } = [];
    public bool HasNoItems => Items.Count == 0;

    // ── 항목 추가 패널 ───────────────────────────────────────────────

    private bool _showAddPanel;
    public bool ShowAddPanel { get => _showAddPanel; set => SetField(ref _showAddPanel, value); }

    private string _newItemLabel = "";
    public string NewItemLabel { get => _newItemLabel; set => SetField(ref _newItemLabel, value); }

    private ItemTypeOption _newItemType = ItemTypeOption.All[0];
    public ItemTypeOption NewItemType
    {
        get => _newItemType;
        set
        {
            if (SetField(ref _newItemType, value))
            {
                // 이름을 아직 안 적었으면 종류 이름으로 채워 준다
                if (string.IsNullOrWhiteSpace(NewItemLabel) || ItemTypeOption.All.Any(o => o.Label == NewItemLabel))
                    NewItemLabel = value.Type == BlockTypes.Date ? "날짜"
                        : value.FixedId != null || value.Type is FieldTypes.DateKorean or FieldTypes.DateSimple
                        ? value.Label.Split(' ')[0] : "";
            }
        }
    }

    public IReadOnlyList<ItemTypeOption> ItemTypes => ItemTypeOption.All;

    private string _newItemSuffix = "";
    public string NewItemSuffix { get => _newItemSuffix; set => SetField(ref _newItemSuffix, value); }

    private string _newItemTime = "";
    public string NewItemTime { get => _newItemTime; set => SetField(ref _newItemTime, value); }

    // ════════════════════════════════════════════════════════════════
    // ── Excel 에서 고르기 (피커) ──────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    private SetupItemViewModel? _pickItem;
    private PickKind _pickKind;
    private string?  _pickColumn;
    /// <summary>사진 구역: 이번 단계에서 [확인]으로 자리를 지정했는지 (아니면 [건너뛰기])</summary>
    private bool _photoRangeAdded;

    public bool IsPicking => _pickItem != null;

    /// <summary>지금 고르는 중인 항목 (카드 강조용)</summary>
    public SetupItemViewModel? PickingItem => _pickItem;

    /// <summary>
    /// 칸·사진 클릭은 선택이 바뀌는 즉시 지정한다(자동).
    /// 여러 개 선택(그림 여러 개, 행 범위)은 다 고른 뒤 [확인]을 눌러야 한다.
    /// </summary>
    public bool PickNeedsConfirm => _pickKind is PickKind.Shapes or PickKind.Rows or PickKind.Lines
                                             or PickKind.PhotoRange or PickKind.SingleRange;

    /// <summary>사진 옆 측정값 칸처럼 없을 수도 있는 단계는 [건너뛰기]</summary>
    public bool PickCanSkip => _pickKind is PickKind.ValueCell or PickKind.PrevCell
                                         or PickKind.PhotoCaption or PickKind.SingleRange or PickKind.SingleCaption
                            || (_pickKind == PickKind.PhotoRange && _pickItem?.PhotoAreas.Count > 0)
                            || (_pickKind == PickKind.Column && _pickItem != null && _pickItem.IsOptionalCol(_pickColumn ?? ""));

    /// <summary>[건너뛰기] 버튼 글자 (사진 구역은 "칸 다 지정함" 처럼 뜻을 바로 보여 준다)</summary>
    public string PickSkipText => _pickKind switch
    {
        PickKind.PhotoRange  => "칸 다 지정함",
        PickKind.SingleRange => "1장일 때 자리 따로 없음",
        PickKind.PhotoCaption or PickKind.SingleCaption => "설명 칸 없음",
        _ => "건너뛰기",
    };

    private string _pickPrompt = "";
    public string PickPrompt { get => _pickPrompt; private set => SetField(ref _pickPrompt, value); }

    /// <summary>View 가 지금 Excel 선택을 "기준"으로 기억하게 하는 신호 (이후 바뀐 선택만 지정)</summary>
    public event EventHandler? PickStarted;

    // ════════════════════════════════════════════════════════════════
    // ── 커맨드 / 이벤트 ───────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    public ICommand BrowseFileCommand     { get; }
    public ICommand SaveCommand           { get; }
    public ICommand AddItemCommand        { get; }
    public ICommand ConfirmAddItemCommand { get; }
    public ICommand CancelAddItemCommand  { get; }
    public ICommand ConfirmPickCommand    { get; }
    public ICommand SkipPickCommand       { get; }
    public ICommand CancelPickCommand     { get; }
    public ICommand ScanPlaceholdersCommand { get; }

    /// <summary>저장 성공 시 View 에 창 닫기를 요청</summary>
    public event EventHandler? SaveSucceeded;

    public DocumentSetupViewModel()
    {
        BrowseFileCommand = new RelayCommand(_ => OnBrowseFile());
        SaveCommand       = new RelayCommand(_ => OnSave());

        AddItemCommand = new RelayCommand(_ =>
        {
            CancelPick();
            NewItemType   = ItemTypeOption.All[0];
            NewItemLabel  = "";
            NewItemSuffix = "";
            NewItemTime   = "";
            ShowAddPanel  = true;
        });
        ConfirmAddItemCommand = new RelayCommand(_ => ConfirmAddItem());
        CancelAddItemCommand  = new RelayCommand(_ => ShowAddPanel = false);

        ConfirmPickCommand = new RelayCommand(_ => ConfirmPick());
        SkipPickCommand    = new RelayCommand(_ =>
        {
            if (_pickKind == PickKind.Column && _pickColumn != null) _skippedCols.Add(_pickColumn);
            if (_pickKind == PickKind.PrevCell) _pickItem?.SetPrevSkipped();
            _photoRangeAdded = false;
            NextPickStep();
        });
        ConvertToBlocksCommand = new RelayCommand(_ => ConvertToBlocks(), _ => HasLegacyItems);
        CancelPickCommand  = new RelayCommand(_ => CancelPick());
        ScanPlaceholdersCommand = new RelayCommand(_ => ScanPlaceholders(showMessage: true));

        Items.CollectionChanged += (_, e) =>
        {
            OnPropertyChanged(nameof(HasNoItems));
            OnPropertyChanged(nameof(HasLegacyItems));
            OnPropertyChanged(nameof(ShowSujangjae));
            if (e.NewItems != null)
                foreach (SetupItemViewModel it in e.NewItems)
                    it.PropertyChanged += (_, pe) =>
                    {
                        if (pe.PropertyName is nameof(SetupItemViewModel.Key) or nameof(SetupItemViewModel.Label)) RefreshKnownKeys();
                        if (pe.PropertyName == nameof(SetupItemViewModel.Summary)) RefreshHighlights();
                    };
            RefreshKnownKeys();
            RefreshHighlights();
        };
    }

    // ════════════════════════════════════════════════════════════════
    // ── 블록 값 이름 / 예전 방식 바꾸기 ───────────────────────────
    // ════════════════════════════════════════════════════════════════

    private List<string>? _otherDocKeys;

    /// <summary>다른 문서들의 블록 이름 + 이 문서 블록 이름 (처음 값·계산·명단 고르기용)</summary>
    public IReadOnlyList<string> KnownKeys()
    {
        _otherDocKeys ??= AppConfigLoader.Load().Documents
            .Where(d => d.Id != _editing?.Id)
            .SelectMany(d => d.Fields.Where(f => BlockTypes.IsBlock(f.Type) && f.Type is not (BlockTypes.Date or BlockTypes.Formula)).Select(BlockTypes.KeyOf))
            .ToList();
        return Items.Where(i => i.HasKey).Select(i => i.EffectiveKey).Concat(_otherDocKeys)
                    .Where(k => k.Length > 0).Distinct().ToList();
    }

    private void RefreshKnownKeys()
    {
        foreach (var i in Items) i.RefreshKnownKeys();
    }

    /// <summary>예전 방식 항목(날씨·장비 표 등 뜻이 코드에 정해진 것)이 있는지</summary>
    public bool HasLegacyItems => Items.Any(i => i.IsLegacy);

    public ICommand ConvertToBlocksCommand { get; }

    /// <summary>
    /// [블록 방식으로 바꾸기]: 예전 항목을 같은 칸·같은 결과의 블록으로 바꾼다 (저장해야 반영).
    /// 바꾼 뒤에는 규칙(값 이름, 처음 값, 계산 등)을 카드에서 바꿀 수 있다.
    /// </summary>
    private void ConvertToBlocks()
    {
        var fields = Items.Select(i => i.ToConfig()).ToList();
        var converted = BlockMigration.Convert(new DocumentConfig { Fields = fields });
        CancelPick();
        Items.Clear();
        foreach (var f in converted) Items.Add(SetupItemViewModel.FromConfig(this, f));
        MessageBox.Show("예전 방식 항목을 블록으로 바꿨습니다. 칸 위치와 결과는 같습니다.\n" +
                        "카드에서 규칙을 확인하고 [저장]을 누르세요.", "블록 방식으로 바꾸기",
                        MessageBoxButton.OK, MessageBoxImage.Information);
    }

    /// <summary>
    /// 지정한 시트가 "적용할 시트" 가 아니면 항목에 그 시트를 기록한다 (그 시트에만 쓰게).
    /// 적용할 시트에서 지정하면 비워 둔다 → 적용할 시트 전부에 쓴다.
    /// </summary>
    private void RecordSheet(SetupItemViewModel item)
    {
        string current = SelectedSheet ?? "";
        bool isTarget = SheetChoices.Any(c => c.IsChecked && c.Name == current);
        item.Sheet = isTarget || current.Length == 0 ? null : current;
    }

    private readonly HashSet<string> _skippedCols = [];

    /// <summary>
    /// 표·명단: 드래그한 행 위의 머리글(장비명·전일·금일·누계 / 직종·성명·작업내용)을 읽어 열 역할을 채운다.
    /// 찾은 게 있으면 못 찾은 "선택" 열은 묻지 않는다 (필수 열만 클릭). 다르면 카드의 열 버튼으로 다시 지정.
    /// </summary>
    private void AutoDetectColumns(SetupItemViewModel item, ExcelSelection sel)
    {
        if (Bridge == null || item.Type is not (BlockTypes.Table or BlockTypes.Roster)) return;
        Dictionary<string, string> found;
        try { found = HeaderDetector.Detect(item.Type, sel.Row, sel.Col, sel.Col + Math.Max(1, sel.ColCount) - 1, Bridge.GetCellText); }
        catch (Exception ex) { AppLog.Error("머리글 인식 실패", ex); return; }

        if (found.Count == 0) { item.AutoNote = ""; return; }
        foreach (var (role, col) in found) item.SetColumn(role, col);
        foreach (var k in item.ColKeys.Where(k => !found.ContainsKey(k) && item.IsOptionalCol(k))) _skippedCols.Add(k);
        item.AutoNote = "머리글에서 찾음: " + string.Join(" · ", item.ColKeys.Where(found.ContainsKey)
                            .Select(k => $"{SetupItemViewModel.ColTitle(k)} {found[k]}"))
                      + "  (다르면 아래 열 버튼을 눌러 다시 지정)";
        if (item.IsTable && item.HasRosterKey) LoadRoleMap(item);
    }

    /// <summary>
    /// 짝짓기 표: 표의 줄 이름(서식의 그 표)과 명단 직종(이 문서·다른 문서의 명단 + 저장된 사람 목록)을 읽는다.
    /// Excel 이 서식을 열고 있어도 읽을 수 있게 읽기 전용 복사로 연다.
    /// </summary>
    public void LoadRoleMap(SetupItemViewModel item)
    {
        if (!HasFile || IsHwpx || !item.IsTable || !item.HasRosterKey || item.StartRow == 0) return;
        try
        {
            string defaultSheet = SheetChoices.FirstOrDefault(c => c.IsChecked)?.Name ?? SelectedSheet ?? "";
            var rows = ReadBlock(_filePath, item.ToConfig(), defaultSheet).Rows.Select(r => r.Name).ToList();

            string rosterKey = item.RosterKey.Trim();
            var jobs = new List<string>();
            var rosterItem = Items.FirstOrDefault(i => i.Type == BlockTypes.Roster && i.EffectiveKey == rosterKey);
            if (rosterItem != null)
                jobs.AddRange(ReadBlock(_filePath, rosterItem.ToConfig(), defaultSheet).People.Select(p => p.Job));
            var cfg = AppConfigLoader.Load();
            foreach (var doc in cfg.Documents.Where(d => d.Id != _editing?.Id))
                foreach (var f in doc.Fields.Where(f => f.Type == BlockTypes.Roster && BlockTypes.KeyOf(f) == rosterKey))
                {
                    string path = AppConfigLoader.ResolvePath(doc.TemplateFile);
                    if (File.Exists(path)) jobs.AddRange(ReadBlock(path, f, doc.TargetSheets().FirstOrDefault() ?? "").People.Select(p => p.Job));
                }
            jobs.AddRange(cfg.People.Select(p => p.Role));
            jobs.AddRange(rows);   // 줄 이름과 같은 직종도 고를 수 있게

            item.SetRoleMapData(rows, jobs);
        }
        catch (Exception ex)
        {
            AppLog.Error("짝짓기 표 읽기 실패", ex);
            Warn($"짝짓기 표를 만들지 못했습니다: {ex.Message}");
        }
    }

    private static Core.Models.BlockValue ReadBlock(string file, FieldConfig f, string sheet)
    {
        var doc = new DocumentConfig { SheetName = sheet, Fields = [f] };
        return BlockEngine.ReadValues(file, doc, sheet).TryGetValue(BlockTypes.KeyOf(f), out var v) ? v : new();
    }

    // ════════════════════════════════════════════════════════════════
    // ── 기존 문서 불러오기 (편집) ─────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    public void LoadExisting(DocumentConfig doc)
    {
        _editing = doc;
        DocumentName      = doc.Name;
        OutputPattern     = doc.OutputPattern;
        SelectedFrequency = FrequencyOption.All.FirstOrDefault(f => f.Value == doc.Frequency) ?? FrequencyOption.All[0];
        FrequencyDays     = doc.FrequencyDays > 0 ? doc.FrequencyDays : 7;
        SelectedWeekday   = WeekdayOption.Of(doc.WeeklyDay);
        UseWeeklyDayDate  = doc.UseWeeklyDayDate;
        IncludesSujangjae = doc.IncludesSujangjae;
        _initialSlotLabels = doc.PhotoSlots.ToDictionary(s => s.Index, s => s.Label);
        _selectedPhotoSheet = string.IsNullOrEmpty(doc.PhotoSheet) ? NoPhotoOption : doc.PhotoSheet;

        foreach (var f in doc.Fields)
            Items.Add(SetupItemViewModel.FromConfig(this, f));

        // 파일 경로를 마지막에 넣어야 시트 목록·적용 시트가 저장값으로 맞춰진다
        _filePath = ResolvePath(doc.TemplateFile);
        OnPropertyChanged(nameof(FilePath));
        OnPropertyChanged(nameof(IsHwpx)); OnPropertyChanged(nameof(IsExcel)); OnPropertyChanged(nameof(HasFile));
        ReloadSheets(doc.TargetSheets());
        OnPropertyChanged(nameof(WindowTitle));
    }

    private static string ResolvePath(string p) =>
        Path.IsPathRooted(p) ? p : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, p);

    // ════════════════════════════════════════════════════════════════
    // ── 파일 ──────────────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    private void OnBrowseFile()
    {
        var dialog = new OpenFileDialog
        {
            Title  = "서식 파일 선택",
            Filter = "서식 파일 (*.xlsx;*.hwpx)|*.xlsx;*.hwpx|Excel (*.xlsx)|*.xlsx|한글 (*.hwpx)|*.hwpx",
        };
        if (dialog.ShowDialog() == true)
            SetNewFile(dialog.FileName);
    }

    public void HandleDroppedFile(string[] files)
    {
        string? f = files.FirstOrDefault(x => x.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                                           || x.EndsWith(".hwpx", StringComparison.OrdinalIgnoreCase));
        if (f != null) SetNewFile(f);
        else if (files.Any(x => x.EndsWith(".hwp", StringComparison.OrdinalIgnoreCase)))
            MessageBox.Show("예전 한글 형식(.hwp)은 읽을 수 없습니다.\n한글에서 [다른 이름으로 저장 → 파일 형식: HWPX]로 저장한 뒤 그 파일을 넣어 주세요.",
                            "한글 파일", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void SetNewFile(string path)
    {
        FilePath = path;

        // 이름·출력 파일명이 비어 있으면 파일 이름으로 채워 준다
        string baseName = Path.GetFileNameWithoutExtension(path);
        string cleaned  = System.Text.RegularExpressions.Regex.Replace(baseName, @"[\s_\-]*\(?\d{2,4}[.\-]\d{1,2}([.\-]\d{1,2})?\)?$", "");
        if (string.IsNullOrWhiteSpace(DocumentName)) DocumentName = cleaned;
        if (string.IsNullOrWhiteSpace(OutputPattern))
            OutputPattern = $"{cleaned}({{yy}}.{{MM}}.{{dd}}){Path.GetExtension(path)}";

        if (IsHwpx) ScanPlaceholders(showMessage: false);
    }

    /// <summary>시트 목록과 "적용할 시트" 칩을 다시 만든다.</summary>
    private void ReloadSheets(List<string>? checkedSheets = null)
    {
        SheetNames.Clear();
        SheetChoices.Clear();

        if (IsExcel && HasFile)
        {
            foreach (var s in ExcelAnalyzer.GetSheetNames(_filePath))
                SheetNames.Add(s);

            var check = checkedSheets?.Where(SheetNames.Contains).ToList();
            if (check == null || check.Count == 0) check = SheetNames.Take(1).ToList();
            foreach (var s in SheetNames)
                SheetChoices.Add(new SheetChoice(s, check.Contains(s)));

            SelectedSheet = check.FirstOrDefault() ?? "";
        }
        else SelectedSheet = "";

        PhotoSheetOptions.Clear();
        PhotoSheetOptions.Add(NoPhotoOption);
        foreach (var s in SheetNames) PhotoSheetOptions.Add(s);
        if (!PhotoSheetOptions.Contains(_selectedPhotoSheet)) _selectedPhotoSheet = NoPhotoOption;
        OnPropertyChanged(nameof(SelectedPhotoSheet));
        ReloadPhotoSlotRows();
    }

    private void ReloadPhotoSlotRows()
    {
        var current = PhotoSlotRows.ToDictionary(r => r.SlotIndex, r => r.Label);
        PhotoSlotRows.Clear();
        if (string.IsNullOrEmpty(PhotoSheet) || !HasFile || IsHwpx) return;

        try
        {
            foreach (var slot in ExcelPhotoWriter.GetPhotoSlotInfos(_filePath, PhotoSheet))
            {
                string label = current.TryGetValue(slot.SlotIndex, out var c) ? c
                             : _initialSlotLabels.TryGetValue(slot.SlotIndex, out var i) ? i : "";
                byte[]? img = ExcelPhotoWriter.GetSlotImageBytes(_filePath, PhotoSheet, slot.SlotIndex);
                PhotoSlotRows.Add(new PhotoSlotLabelRow(slot.SlotIndex, slot.WidthPx, slot.HeightPx, label, img));
            }
        }
        catch { /* 파일 접근 실패 시 비움 */ }
    }

    /// <summary>한글 서식에 적힌 {{...}} 자리를 찾아 아직 없는 것은 "매일 입력" 항목으로 추가한다.</summary>
    private void ScanPlaceholders(bool showMessage)
    {
        if (!IsHwpx || !HasFile) return;
        List<string> found;
        try { found = HwpxWriter.FindPlaceholders(_filePath); }
        catch (Exception ex)
        {
            MessageBox.Show($"한글 파일을 읽을 수 없습니다.\n{ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        int added = 0;
        foreach (var ph in found.Where(ph => Items.All(i => i.Placeholder != ph)))
        {
            string label = ph.Trim('{', '}').Trim();
            string type = GuessTypeForLabel(label);
            var item = new SetupItemViewModel(this, MakeId(label), label, type) { Placeholder = ph };
            if (type == BlockTypes.Date) item.DateFormat = "{yyyy}.{MM}.{dd}";
            Items.Add(item);
            added++;
        }

        if (showMessage || found.Count == 0)
            MessageBox.Show(found.Count == 0
                ? "서식에서 {{ }} 자리를 찾지 못했습니다.\n\n한글에서 채울 자리에 {{날짜}}, {{날씨}} 처럼 적어 두고 hwpx 로 저장한 뒤 다시 불러오세요.\n(자리 글자는 한 번에 입력하고, 중간에 글꼴을 바꾸지 마세요)"
                : $"자리 {found.Count}개를 찾았고, 새 항목 {added}개를 추가했습니다.\n종류가 다르면 항목을 지우고 [+ 항목 추가]로 다시 만드세요.",
                "한글 서식", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    /// <summary>자리 이름으로 종류 추측: 날짜·일자 → 날짜, 나머지 → 값(매일 입력). 카드에서 바꿀 수 있다.</summary>
    private static string GuessTypeForLabel(string label) =>
        label.Contains("날짜") || label.Contains("일자") ? BlockTypes.Date : BlockTypes.Value;

    // ════════════════════════════════════════════════════════════════
    // ── 항목 추가 / 삭제 ──────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    private void ConfirmAddItem()
    {
        var type = NewItemType;
        string label = string.IsNullOrWhiteSpace(NewItemLabel) ? type.Label.Split(' ')[0] : NewItemLabel.Trim();

        if (type.FixedId != null && Items.Any(i => i.Id == type.FixedId))
        {
            MessageBox.Show($"'{type.Label}' 항목은 이미 있습니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        string id = type.FixedId ?? MakeId(label);
        var item = new SetupItemViewModel(this, id, label, type.Type)
        {
            Suffix = type.AsksSuffix && !string.IsNullOrWhiteSpace(NewItemSuffix) ? NewItemSuffix.Trim() : null,
            Time   = type.AsksTime && !string.IsNullOrWhiteSpace(NewItemTime) ? NormalizeTime(NewItemTime) : null,
            Placeholder = IsHwpx ? "{{" + label + "}}" : null,
        };
        Items.Add(item);
        ShowAddPanel = false;

        // 바로 Excel 에서 고르기 시작
        if (IsExcel && HasFile) StartPick(item, item.MainPickKind);
    }

    /// <summary>"10" / "10시" / "1000" / "10:00" → "10:00"</summary>
    private static string NormalizeTime(string s)
    {
        string digits = new(s.Where(char.IsDigit).ToArray());
        if (TimeSpan.TryParse(s, out var t) && s.Contains(':')) return $"{t:hh\\:mm}";
        return digits.Length switch
        {
            1 or 2 => $"{int.Parse(digits):00}:00",
            3 or 4 => $"{int.Parse(digits[..^2]):00}:{digits[^2..]}",
            _ => s.Trim(),
        };
    }

    /// <summary>이름으로 고유한 항목 ID 를 만든다 (같은 이름이 있으면 번호를 붙인다)</summary>
    private string MakeId(string label)
    {
        string baseId = new string(label.Trim().Select(c => char.IsLetterOrDigit(c) ? c : '_').ToArray()).ToLowerInvariant();
        if (baseId.Length == 0) baseId = "item";
        string id = baseId;
        for (int n = 2; Items.Any(i => i.Id == id) || id is "weather" or "today_work" or "tomorrow_work"; n++)
            id = $"{baseId}_{n}";
        return id;
    }

    public void RemoveItem(SetupItemViewModel item)
    {
        if (_pickItem == item) CancelPick();
        if (EditingItem == item) EditingItem = null;
        Items.Remove(item);
    }

    public void RevealItem(SetupItemViewModel item)
    {
        if (IsPicking || Bridge == null) return;
        Bridge.Reveal(item.Type == FieldTypes.Photo ? item.Cell
                      : item.Type is FieldTypes.Multiline or BlockTypes.Lines ? item.Cells.FirstOrDefault()
                      : item.IsPhotos ? item.PhotoAreas.FirstOrDefault()?.Range.Split(':')[0]
                      : item.StartRow > 0 ? $"{item.ColKeys.Select(item.ColOf).FirstOrDefault(c => c != null) ?? "A"}{item.StartRow}"
                      : item.Cell ?? item.TotalCell,
                      item.Type == FieldTypes.ShapePosition ? item.ShapeIndexes : null);
    }

    // ════════════════════════════════════════════════════════════════
    // ── 피커 ──────────────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    public void StartPick(SetupItemViewModel item, PickKind kind, string? columnKey = null)
    {
        if (IsHwpx || Bridge == null)
        {
            MessageBox.Show("왼쪽에 Excel 서식이 열려 있어야 칸을 고를 수 있습니다.", "알림",
                            MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        _pickItem   = item;
        _pickKind   = kind;
        _pickColumn = columnKey;

        PickPrompt = kind switch
        {
            PickKind.Cell      => $"「{item.Label}」 값을 쓸 칸을 클릭하세요.",
            PickKind.Picture   => $"「{item.Label}」 사진이 들어갈 그림(사진 칸)을 클릭하세요.",
            PickKind.ValueCell => $"「{item.Label}」 측정값을 쓸 칸을 클릭하세요. 없으면 [건너뛰기].",
            PickKind.Shapes    => $"「{item.Label}」 함께 옮길 그림·글상자를 클릭하세요. 여러 개면 Ctrl 을 누른 채 클릭한 뒤 [확인].",
            PickKind.Rows      => $"「{item.Label}」 표에서 값이 들어가는 행들을 위에서 아래로 드래그한 뒤 [확인].",
            PickKind.Lines     => $"「{item.Label}」 한 줄씩 들어갈 칸들을 위에서 아래로 드래그한 뒤 [확인].",
            PickKind.Column    => $"「{item.Label}」 {SetupItemViewModel.ColTitle(columnKey ?? "")} 열의 아무 칸이나 클릭하세요."
                                  + (item.IsOptionalCol(columnKey ?? "") ? " 없으면 [건너뛰기]." : ""),
            PickKind.TotalCell => $"「{item.Label}」 누계 칸을 클릭하세요 (누계 = 어제 누계 + 오늘).",
            PickKind.PrevCell  => $"「{item.Label}」 전일(어제까지 누계) 칸이 따로 있으면 클릭하세요. 없으면 [건너뛰기].",
            PickKind.PhotoRange => $"「{item.Label}」 사진 칸 {item.PhotoAreas.Count + 1} 자리를 드래그(또는 그 자리 그림 클릭)한 뒤 [확인]."
                                   + (item.PhotoAreas.Count > 0 ? " 칸이 더 없으면 [칸 다 지정함]." : ""),
            PickKind.PhotoCaption => $"「{item.Label}」 사진 칸 {item.PhotoAreas.Count} 의 설명 칸을 클릭하세요. 없으면 [설명 칸 없음].",
            PickKind.SingleRange  => $"「{item.Label}」 사진이 1장인 날 넣을 자리를 드래그(예: 두 칸을 합친 자리)한 뒤 [확인]. 따로 없으면 칸 1 에 넣습니다.",
            PickKind.SingleCaption => $"「{item.Label}」 1장일 때 설명 칸을 클릭하세요. 없으면 [설명 칸 없음].",
            PickKind.FormulaYesterday => $"「{item.Label}」 식에 넣을 칸을 클릭하세요. 어제 파일의 그 칸 값을 씁니다.",
            PickKind.FormulaCell      => $"「{item.Label}」 식에 넣을 칸을 클릭하세요. 오늘 파일의 그 칸 값(다른 항목을 다 쓴 뒤)을 씁니다.",
            _ => "",
        };

        NotifyPick();
        PickStarted?.Invoke(this, EventArgs.Empty);
    }

    public void CancelPick()
    {
        if (_pickItem == null) return;
        _pickItem = null;
        PickPrompt = "";
        NotifyPick();
    }

    private void NotifyPick()
    {
        OnPropertyChanged(nameof(IsPicking));
        OnPropertyChanged(nameof(PickingItem));
        OnPropertyChanged(nameof(PickNeedsConfirm));
        OnPropertyChanged(nameof(PickCanSkip));
        OnPropertyChanged(nameof(PickSkipText));
    }

    /// <summary>
    /// View 가 Excel 선택이 바뀔 때마다 호출한다 (고르는 중일 때만).
    /// 클릭 한 번으로 끝나는 단계(칸·사진·열)는 여기서 바로 지정한다.
    /// </summary>
    public void OnExcelSelectionChanged(ExcelSelection sel)
    {
        var item = _pickItem;
        if (item == null) return;

        switch (_pickKind)
        {
            case PickKind.Cell when !sel.IsShapes:
                item.Cell = sel.ActiveCell;
                RecordSheet(item);
                NextPickStep();
                break;

            case PickKind.TotalCell when !sel.IsShapes:
                item.TotalCell = sel.ActiveCell;
                RecordSheet(item);
                NextPickStep();
                break;

            case PickKind.PrevCell when !sel.IsShapes:
                item.PrevCell = sel.ActiveCell;
                NextPickStep();
                break;

            case PickKind.ValueCell when !sel.IsShapes:
                item.ValueCell = sel.ActiveCell;
                NextPickStep();
                break;

            case PickKind.Column when !sel.IsShapes:
                item.SetColumn(_pickColumn!, new string(sel.ActiveCell.TakeWhile(char.IsLetter).ToArray()));
                NextPickStep();
                break;

            case PickKind.FormulaYesterday when !sel.IsShapes:
                item.AppendToFormula($"[어제 {sel.ActiveCell}]");
                CancelPick();
                break;

            case PickKind.FormulaCell when !sel.IsShapes:
                item.AppendToFormula($"[{sel.ActiveCell}]");
                CancelPick();
                break;

            case PickKind.PhotoCaption when !sel.IsShapes:
                item.SetLastCaption(sel.ActiveCell);
                NextPickStep();
                break;

            case PickKind.SingleCaption when !sel.IsShapes:
                item.SetSingleCaption(sel.ActiveCell);
                NextPickStep();
                break;

            case PickKind.Picture when sel.IsShapes && sel.ShapeNames.Count == 1:
                item.Cell = sel.ActiveCell;
                RecordSheet(item);
                NextPickStep();
                break;
        }
    }

    /// <summary>[확인]: 여러 개를 고르는 단계(그림 여러 개, 행 범위)를 현재 선택으로 마친다.</summary>
    private void ConfirmPick()
    {
        var sel = Bridge?.GetSelection();
        if (sel != null) ConfirmPickWith(sel);
    }

    private void ConfirmPickWith(ExcelSelection sel)
    {
        var item = _pickItem;
        if (item == null) return;

        switch (_pickKind)
        {
            case PickKind.Shapes:
                if (!sel.IsShapes) { Warn("그림·글상자를 먼저 클릭해 선택하세요."); return; }
                item.SetShapes(sel.ShapeNames, sel.ShapeIndexes);
                break;

            case PickKind.Rows:
                if (sel.IsShapes) { Warn("표의 칸들을 드래그해 선택하세요."); return; }
                item.SetRows(sel.Row, sel.Row + sel.RowCount - 1);
                RecordSheet(item);
                _skippedCols.Clear();
                AutoDetectColumns(item, sel);
                break;

            case PickKind.Lines:
                if (sel.IsShapes) { Warn("칸들을 드래그해 선택하세요."); return; }
                string col = new(sel.ActiveCell.TakeWhile(char.IsLetter).ToArray());
                item.SetLines(Enumerable.Range(sel.Row, sel.RowCount).Select(r => $"{col}{r}").ToList());
                RecordSheet(item);
                break;

            case PickKind.PhotoRange:
            case PickKind.SingleRange:
                // 칸을 드래그했거나, 그 자리에 있는 그림을 클릭했으면 그림이 덮는 칸 범위
                if (sel.ColCount * sel.RowCount < 2) { Warn("사진이 들어갈 칸들을 드래그하거나, 그 자리의 그림을 클릭하세요."); return; }
                string range = $"{HeaderDetector.ColLetter(sel.Col)}{sel.Row}:" +
                               $"{HeaderDetector.ColLetter(sel.Col + sel.ColCount - 1)}{sel.Row + sel.RowCount - 1}";
                if (_pickKind == PickKind.PhotoRange) item.AddPhotoArea(range); else item.SetSingleArea(range);
                RecordSheet(item);
                _photoRangeAdded = true;
                break;

            default:
                return;
        }
        NextPickStep();
    }

    /// <summary>
    /// 다음 단계로: 사진 → 측정값 칸, 표 행 → 각 열, 그림 묶음 → (위치 저장 안내).
    /// 더 할 게 없으면 고르기를 끝낸다.
    /// </summary>
    private void NextPickStep()
    {
        var item = _pickItem;
        if (item == null) return;

        // 사진 칸을 고른 직후 → 측정값 칸
        if (_pickKind == PickKind.Picture)
        {
            StartPick(item, PickKind.ValueCell);
            return;
        }

        // 표·명단: 행 범위 → 아직 안 정한 열을 순서대로 (건너뛴 선택 열은 다시 묻지 않음)
        if (_pickKind is PickKind.Rows or PickKind.Column)
        {
            string? nextCol = item.ColKeys.FirstOrDefault(k => item.ColOf(k) == null && !_skippedCols.Contains(k));
            if (nextCol != null) { StartPick(item, PickKind.Column, nextCol); return; }
        }

        // 사진 구역: 칸 자리 → 설명 칸 → 다음 칸 자리 … [칸 다 지정함] → (2칸 이상이면) 1장일 때 자리 → 그 설명 칸
        if (item.IsPhotos)
        {
            switch (_pickKind)
            {
                case PickKind.PhotoRange when _photoRangeAdded:
                    StartPick(item, PickKind.PhotoCaption); return;
                case PickKind.PhotoCaption:
                    StartPick(item, PickKind.PhotoRange); return;
                case PickKind.PhotoRange when item.PhotoAreas.Count >= 2 && item.SinglePhotoArea == null:
                    StartPick(item, PickKind.SingleRange); return;
                case PickKind.SingleRange when item.SinglePhotoArea != null && _photoRangeAdded:
                    StartPick(item, PickKind.SingleCaption); return;
            }
        }

        // 숫자 누계: 오늘 칸 → 누계 칸 → 전일 칸(선택)
        // (전일 칸은 자동으로 묻지 않는다: 칸을 먼저 고르는 방식에서는 다음에 클릭한 칸이 전일로 잡혀 헷갈림.
        //  필요하면 카드의 [전일 칸 (선택)])
        if (item.IsCounter && _pickKind == PickKind.Cell && item.TotalCell == null) { StartPick(item, PickKind.TotalCell); return; }

        CancelPick();

        if (_pickKind == PickKind.Shapes && item.Presets.Count == 0)
            MessageBox.Show("이제 Excel 에서 그림을 원하는 자리로 끌어다 놓고,\n항목 카드의 [지금 위치 저장]을 누르세요.\n위치를 여러 개 저장해 두면 매일 그중에서 고를 수 있습니다.",
                            "위치 저장 안내", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    // ════════════════════════════════════════════════════════════════
    // ── 칸 먼저 고르기: Excel 에서 칸을 고르면 "이 칸은?" 선택지 ────
    // ════════════════════════════════════════════════════════════════

    private ExcelSelection? _freeSel;

    /// <summary>고른 칸·범위·그림에 맞는 선택지 (위쪽 "이 칸은?" 상자)</summary>
    public ObservableCollection<CellChoice> CellChoices { get; } = [];
    public bool HasCellChoices => CellChoices.Count > 0;

    private string _selectionText = "";
    /// <summary>예: "선택한 칸: H20" / "선택한 범위: A9:A17 (9줄)"</summary>
    public string SelectionText { get => _selectionText; private set => SetField(ref _selectionText, value); }

    private string _choiceQuestion = "";
    /// <summary>예: "이 칸에는 무엇이 들어가나요?"</summary>
    public string ChoiceQuestion { get => _choiceQuestion; private set => SetField(ref _choiceQuestion, value); }

    private SetupItemViewModel? _editingItem;
    /// <summary>방금 만든 항목: 위쪽 상자에서 바로 이름·형식 등을 고친다 (아래 목록까지 내려가지 않게)</summary>
    public SetupItemViewModel? EditingItem
    {
        get => _editingItem;
        private set { if (SetField(ref _editingItem, value)) OnPropertyChanged(nameof(HasEditingItem)); }
    }
    public bool HasEditingItem => _editingItem != null;
    public ICommand CloseEditingCommand => new RelayCommand(_ => EditingItem = null);

    private string _selectionOwner = "";
    /// <summary>이미 그 칸을 쓰는 항목이 있으면 안내 (예: "이 칸은 '날씨' 항목이 씁니다")</summary>
    public string SelectionOwner { get => _selectionOwner; private set { if (SetField(ref _selectionOwner, value)) OnPropertyChanged(nameof(HasSelectionOwner)); } }
    public bool HasSelectionOwner => _selectionOwner.Length > 0;

    public ICommand CloseCellChoicesCommand => new RelayCommand(_ => ClearCellChoices());

    /// <summary>새 항목을 만들어 목록에 넣었을 때 (View 가 그 카드로 스크롤)</summary>
    public event EventHandler<SetupItemViewModel>? ItemCreated;

    /// <summary>View 가 고르는 중이 아닐 때 Excel 선택이 바뀌면 부른다.</summary>
    public void OnFreeSelection(ExcelSelection sel)
    {
        if (IsPicking || IsHwpx) return;
        _freeSel = sel;
        EditingItem = null;
        CellChoices.Clear();

        // 종류 이름 대신 "무엇이 들어가나요?" 에 대한 답으로 고른다 (초보자도 고를 수 있게)
        if (sel.IsShapes)
        {
            SelectionText  = $"선택한 그림: {string.Join(", ", sel.ShapeNames)}";
            ChoiceQuestion = "이 그림은 매일 어떻게 하나요?";
            Add("📷", "그날 찍은 사진으로 바꿔요", BlockTypes.Photos, "이 그림 자리에 매일 그날 사진을 넣습니다. 사진 칸이 여러 개면 이어서 지정합니다.");
            if (sel.ShapeNames.Count == 1)
                Add("📷", "사진으로 바꾸고 옆 칸에 측정값도 써요", FieldTypes.Photo, "소음 측정처럼 사진 + 숫자 한 칸.");
            Add("📍", "위치만 옮겨요 (장비·측정기 표시)", FieldTypes.ShapePosition, "이 그림·글상자를 저장해 둔 위치 중 하나로 옮깁니다.");
        }
        else if (sel.IsOneCell)
        {
            SelectionText  = $"선택한 칸: {sel.ActiveCell}";
            ChoiceQuestion = "이 칸에는 무엇이 들어가나요?";
            Add("📅", "날짜", BlockTypes.Date, "만드는 날 날짜를 원하는 모양으로 씁니다 (2026.11.20 / 2026년 11월 20일 금요일 …).");
            Add("✏️", "매일 다른 값 (날씨·인원·측정값 등)", BlockTypes.Value, "메인 화면에서 매일 입력합니다. 어제 값이 미리 채워져 있어 바뀐 것만 고치면 됩니다.");
            Add("📌", "늘 같은 값 (현장명·참석율 등)", BlockTypes.Value, "매일 같은 값을 씁니다.", i => i.ValueSource = SetupItemViewModel.ValueSourceList[2]);
            Add("➕", "어제까지 누계에 오늘을 더한 값", BlockTypes.Counter, "이 칸에는 오늘 숫자, 이어서 클릭할 칸에는 누계(어제 누계 + 오늘)를 씁니다.");
            Add("🧮", "다른 항목들의 합계·이름 모음", BlockTypes.Calc, "인원 표의 금일 합계, 명단 인원 수, 투입 장비 이름 나열 등.");
            Add("🧾", "계산식 직접 만들기 (고급)", BlockTypes.Formula, "어제 칸·오늘 칸·입력값을 더하고 빼거나 글자로 이어붙입니다. 예) [어제 H20] + [입력 장비]");
            if (SheetChoices.Count(c => c.IsChecked) > 1)
                Add("✏️", "시트마다 다른 값", FieldTypes.Input, "같은 모양 시트마다 따로 입력합니다 (측정 장소별 값 등).");
        }
        else
        {
            SelectionText  = $"선택한 범위: {RangeOf(sel)} ({sel.RowCount}줄 × {sel.ColCount}열)";
            ChoiceQuestion = "이 범위에는 무엇이 들어가나요?";
            Add("📝", "여러 줄 글 (작업내용 등)", BlockTypes.Lines, "한 줄에 하나씩 입력합니다. 어제 내용에서 시작하게 할 수도 있습니다.");
            if (sel.ColCount > 1)
            {
                Add("📊", "이름별 숫자 표 (장비·인원 투입현황)", BlockTypes.Table, "줄마다 이름과 전일·금일·누계. 머리글을 읽어 열을 찾고, 못 찾은 열만 클릭합니다.");
                Add("👥", "사람 명단 (한 줄에 한 사람)", BlockTypes.Roster, "직종·성명·작업내용. 메인 화면에서 오늘 나온 사람만 체크합니다.");
            }
            Add("📷", "그날 찍은 사진", BlockTypes.Photos, "이 범위에 사진을 꽉 채워 넣습니다. 사진 칸이 여러 개면 이어서 지정합니다.");
        }

        string? owner = Items.FirstOrDefault(i => i.UsesCell(sel.ActiveCell))?.Label;
        SelectionOwner = owner == null ? "" : $"이 칸은 이미 '{owner}' 항목이 씁니다. 새로 만들면 둘 다 씁니다.";
        OnPropertyChanged(nameof(HasCellChoices));

        void Add(string icon, string label, string type, string hint, Action<SetupItemViewModel>? preset = null) =>
            CellChoices.Add(new CellChoice(icon, label, hint, new RelayCommand(_ => CreateFromSelection(type, label, preset))));
    }

    /// <summary>미리보기용: 지금 서식에서 그 칸(예: "H20")에 보이는 글자</summary>
    public string CellText(string address)
    {
        if (Bridge == null || string.IsNullOrWhiteSpace(address)) return "";
        try
        {
            var (col, row) = CellHelper.ParseAddress(address.Trim().ToUpperInvariant());
            return Bridge.GetCellText((int)row, (int)CellHelper.ColLetterToIndex(col));
        }
        catch { return ""; }
    }

    /// <summary>
    /// 이미 정한 칸을 Excel 에 색으로 칠한다. 지금 보이는 시트에 쓰는 항목만
    /// (시트를 따로 정한 항목은 그 시트일 때, 아니면 "적용할 시트" 일 때).
    /// </summary>
    public void RefreshHighlights()
    {
        if (Bridge == null || IsHwpx) return;
        string sheet = SelectedSheet ?? "";
        bool isTarget = SheetChoices.Any(c => c.IsChecked && c.Name == sheet);
        var ranges = Items.Where(i => i.HasOwnSheet ? i.Sheet == sheet : isTarget)
                          .SelectMany(i => i.HighlightRanges()).Distinct().ToList();
        try { Bridge.Highlight(ranges); }
        catch (Exception ex) { AppLog.Error("칸 색 표시 실패", ex); }
    }

    public void ClearCellChoices()
    {
        _freeSel = null;
        CellChoices.Clear();
        SelectionOwner = "";
        OnPropertyChanged(nameof(HasCellChoices));
    }

    private static string RangeOf(ExcelSelection sel) =>
        $"{HeaderDetector.ColLetter(sel.Col)}{sel.Row}:{HeaderDetector.ColLetter(sel.Col + sel.ColCount - 1)}{sel.Row + sel.RowCount - 1}";

    /// <summary>고른 칸으로 항목을 만들고, 그 종류에 남은 단계(누계 칸, 표의 열, 사진 설명 칸 등)를 이어서 묻는다.</summary>
    private void CreateFromSelection(string type, string choiceLabel, Action<SetupItemViewModel>? preset)
    {
        var sel = _freeSel;
        if (sel == null) return;
        ClearCellChoices();

        string label = UniqueLabel(GuessLabel(sel) ?? DefaultLabel(type, choiceLabel));
        var item = new SetupItemViewModel(this, MakeId(label), label, type);
        preset?.Invoke(item);
        Items.Add(item);
        ItemCreated?.Invoke(this, item);
        EditingItem = item;   // 위쪽 상자에서 바로 고치게

        // 고른 선택을 그대로 첫 단계로 쓴다 (다시 클릭하지 않게)
        StartPick(item, item.MainPickKind);
        if (PickNeedsConfirm) ConfirmPickWith(sel);
        else OnExcelSelectionChanged(sel);
    }

    /// <summary>
    /// 칸 이름 짐작: 왼쪽 칸(3칸까지) 글자, 없으면 위 칸 글자. 범위면 첫 칸 위 머리글.
    /// 서식마다 다른 이름을 코드에 두지 않고 서식에 적힌 글자를 쓴다. 카드에서 고칠 수 있다.
    /// </summary>
    private string? GuessLabel(ExcelSelection sel)
    {
        if (Bridge == null || sel.IsShapes) return null;
        var tries = new List<(int Row, int Col)>();
        if (sel.IsOneCell) for (int c = sel.Col - 1; c >= Math.Max(1, sel.Col - 3); c--) tries.Add((sel.Row, c));
        if (sel.Row > 1) tries.Add((sel.Row - 1, sel.Col));
        foreach (var (r, c) in tries)
        {
            string t;
            try { t = Bridge.GetCellText(r, c); } catch { continue; }
            t = new string(t.Replace("\n", " ").Where(ch => !char.IsControl(ch)).ToArray()).Trim().TrimEnd(':', '：').Trim();
            if (t.Length == 0 || BlockTypes.ParseNumber(t) != null) continue;
            return t.Length > 20 ? t[..20] : t;
        }
        return null;
    }

    private static string DefaultLabel(string type, string choiceLabel) => type switch
    {
        BlockTypes.Date    => "날짜",
        BlockTypes.Photos  => "사진",
        BlockTypes.Formula => "계산",
        _                  => choiceLabel.Split(' ')[0],
    };

    /// <summary>같은 문서에 같은 이름이 있으면 번호를 붙인다 (같은 이름 = 같은 입력이 되므로)</summary>
    private string UniqueLabel(string label)
    {
        string result = label;
        for (int n = 2; Items.Any(i => i.Label.Trim() == result); n++) result = $"{label} {n}";
        return result;
    }

    // ── 그림 위치 저장 / 보기 ────────────────────────────────────────

    public void SavePreset(SetupItemViewModel item)
    {
        if (Bridge == null || item.ShapeIndexes.Count == 0) return;

        var placements = Bridge.ReadPlacements(item.ShapeIndexes);
        if (placements.Count == 0) { Warn("Excel 에서 그림을 찾지 못했습니다. 지금 보이는 시트에 그 그림이 있는지 확인하세요."); return; }

        string name = string.IsNullOrWhiteSpace(item.NewPresetName) ? $"위치 {item.Presets.Count + 1}" : item.NewPresetName.Trim();
        var existing = item.Presets.FirstOrDefault(p => p.Name == name);
        if (existing != null) item.RemovePreset(existing);   // 같은 이름이면 덮어쓰기

        item.AddPreset(new ShapePreset { Name = name, Shapes = placements });
        item.NewPresetName = "";
    }

    public void ShowPreset(ShapePreset preset) => Bridge?.ApplyPlacements(preset.Shapes);

    private static void Warn(string msg) =>
        MessageBox.Show(msg, "알림", MessageBoxButton.OK, MessageBoxImage.Information);

    // ════════════════════════════════════════════════════════════════
    // ── 저장 ──────────────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    private void OnSave()
    {
        if (string.IsNullOrWhiteSpace(DocumentName)) { Warn("문서 이름을 입력해 주세요."); return; }
        if (!HasFile) { Warn("서식 파일(.xlsx / .hwpx)을 선택해 주세요."); return; }
        if (string.IsNullOrWhiteSpace(OutputPattern)) { Warn("출력 파일명을 입력해 주세요."); return; }

        var sheets = SheetChoices.Where(s => s.IsChecked).Select(s => s.Name).ToList();
        if (IsExcel && sheets.Count == 0) { Warn("적용할 시트를 하나 이상 골라 주세요."); return; }

        var incomplete = Items.Where(i => !i.IsComplete).Select(i => i.Label).ToList();
        if (incomplete.Count > 0
            && MessageBox.Show($"아직 지정하지 않은 항목이 있습니다:\n{string.Join(", ", incomplete)}\n\n그래도 저장할까요? (지정 안 된 항목은 채우지 않습니다)",
                               "확인", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            return;

        // 출력 파일 확장자를 서식과 맞춘다
        string ext = Path.GetExtension(_filePath);
        if (!OutputPattern.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
            OutputPattern = Path.ChangeExtension(OutputPattern, ext);

        try
        {
            var config = AppConfigLoader.Load();

            string id = _editing?.Id ?? MakeDocId(DocumentName, config);
            var doc = BuildDocument(id, AppConfigLoader.ToStoredTemplatePath(_filePath));

            int idx = config.Documents.FindIndex(d => d.Id == id);
            if (idx >= 0) config.Documents[idx] = doc;
            else config.Documents.Add(doc);
            AppConfigLoader.Save(config);

            SaveSucceeded?.Invoke(this, EventArgs.Empty);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"저장 중 오류가 발생했습니다.\n{ex.Message}", "저장 오류",
                            MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>지금 화면의 설정으로 문서 설정을 만든다 (저장·시험 생성이 같이 쓴다)</summary>
    private DocumentConfig BuildDocument(string id, string templateFile)
    {
        var sheets = SheetChoices.Where(s => s.IsChecked).Select(s => s.Name).ToList();
        return new DocumentConfig
        {
            Id                = id,
            Name              = DocumentName.Trim(),
            TemplateFile      = templateFile,
            OutputPattern     = OutputPattern.Trim(),
            SheetName         = sheets.FirstOrDefault() ?? "",
            SheetNames        = sheets.Count > 1 ? sheets : [],
            Frequency         = SelectedFrequency.Value,
            FrequencyDays     = IsCustomFrequency ? FrequencyDays : 0,
            WeeklyDay         = SelectedWeekday.Value,
            UseWeeklyDayDate  = IsWeeklyFrequency && UseWeeklyDayDate,
            IncludesSujangjae = IncludesSujangjae,
            PhotoSheet        = PhotoSheet,
            PhotoSlots        = PhotoSlotRows.Select(r => r.ToConfig()).ToList(),
            Fields            = Items.Select(i => i.ToConfig()).ToList(),
            // 이 화면에서 다루지 않는 값은 기존 설정을 그대로 유지
            SaveMode          = _editing?.SaveMode ?? "create",
            PhotoFolder       = _editing?.PhotoFolder ?? "",
            UsePreviousAsBase = _editing?.UsePreviousAsBase ?? true,
        };
    }

    // ── 시험 생성 ────────────────────────────────────────────────────

    public ICommand TestGenerateCommand => new RelayCommand(_ => TestGenerate());

    /// <summary>
    /// [🧪 시험 생성]: 저장하지 않고 지금 설정으로 오늘 파일을 임시 폴더에 만들어 본다.
    /// 입력값은 서식에서 읽은 처음 값(매일 입력 칸은 서식 글자, 누계는 0 더하기)이라 칸 위치·형식 확인용.
    /// 실제 출력 폴더·원본 서식은 건드리지 않는다.
    /// </summary>
    private void TestGenerate()
    {
        if (!HasFile || IsHwpx) { Warn("엑셀 서식이 열려 있어야 시험 생성을 할 수 있습니다."); return; }
        if (string.IsNullOrWhiteSpace(OutputPattern)) OutputPattern = "{yy}.{MM}.{dd} 시험.xlsx";
        try
        {
            var doc = BuildDocument("test", _filePath);
            var notes = new List<string>();
            string path = TrialGenerator.Generate(doc, _filePath, TrialGenerator.NewTempFolder(), DateTime.Today, notes);

            string msg = $"시험 파일을 만들었습니다 (저장하지 않음, 임시 폴더).\n{path}\n\n" +
                         "입력 칸은 서식에 적힌 글자를 그대로 넣었습니다. 칸 위치·날짜 형식·식 결과를 확인하세요.";
            if (notes.Count > 0) msg += "\n\n참고:\n· " + string.Join("\n· ", notes.Take(8));
            if (MessageBox.Show(msg + "\n\n지금 Excel 로 열까요?", "시험 생성", MessageBoxButton.YesNo,
                                MessageBoxImage.Information) == MessageBoxResult.Yes)
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            AppLog.Error("시험 생성 실패", ex);
            MessageBox.Show($"시험 생성 중 오류가 났습니다.\n{ex.Message}", "시험 생성", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private static string MakeDocId(string name, AppConfig config)
    {
        string baseId = name.Trim().Replace(" ", "_").Replace("/", "_").ToLowerInvariant();
        string id = baseId;
        for (int n = 2; config.Documents.Any(d => d.Id == id); n++) id = $"{baseId}_{n}";
        return id;
    }
}

// ════════════════════════════════════════════════════════════════
// ── FrequencyOption: 주기 선택 항목 ──────────────────────────
// ════════════════════════════════════════════════════════════════

public class FrequencyOption
{
    public string Label { get; }
    public string Value { get; }

    public FrequencyOption(string label, string value) { Label = label; Value = value; }

    public static readonly IReadOnlyList<FrequencyOption> All =
    [
        new("매일",        "daily"),
        new("매주",        "weekly"),
        new("2주마다",     "biweekly"),
        new("매월",        "monthly"),
        new("사용자 지정", "custom"),
    ];
}

// ── WeekdayOption: 매주 주기의 요일 선택 항목 ─────────────────────
public class WeekdayOption
{
    public string    Label { get; }
    public DayOfWeek Value { get; }

    public WeekdayOption(string label, DayOfWeek value) { Label = label; Value = value; }

    /// <summary>월요일부터 일요일까지 (현장 주간 순서)</summary>
    public static readonly IReadOnlyList<WeekdayOption> All =
    [
        new("월요일", DayOfWeek.Monday), new("화요일", DayOfWeek.Tuesday), new("수요일", DayOfWeek.Wednesday),
        new("목요일", DayOfWeek.Thursday), new("금요일", DayOfWeek.Friday), new("토요일", DayOfWeek.Saturday),
        new("일요일", DayOfWeek.Sunday),
    ];

    public static WeekdayOption Of(DayOfWeek d) => All.First(o => o.Value == d);
}
