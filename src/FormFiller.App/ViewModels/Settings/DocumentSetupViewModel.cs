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
        set => SetField(ref _selectedSheet, value);
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
        set { if (SetField(ref _selectedFrequency, value)) OnPropertyChanged(nameof(IsCustomFrequency)); }
    }

    public bool IsCustomFrequency => _selectedFrequency?.Value == "custom";

    private int _frequencyDays = 7;
    public int FrequencyDays { get => _frequencyDays; set => SetField(ref _frequencyDays, value); }

    private bool _includesSujangjae;
    public bool IncludesSujangjae { get => _includesSujangjae; set => SetField(ref _includesSujangjae, value); }

    // ── 사진대지 슬롯 (예전 방식: 시트의 그림을 순서대로 교체) ──────────

    private const string NoPhotoOption = "없음";
    public ObservableCollection<string> PhotoSheetOptions { get; } = [NoPhotoOption];

    private string _selectedPhotoSheet = NoPhotoOption;
    public string SelectedPhotoSheet
    {
        get => _selectedPhotoSheet;
        set { if (SetField(ref _selectedPhotoSheet, value)) ReloadPhotoSlotRows(); }
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
                    NewItemLabel = value.FixedId != null || value.Type is FieldTypes.DateKorean or FieldTypes.DateSimple
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

    public bool IsPicking => _pickItem != null;

    /// <summary>지금 고르는 중인 항목 (카드 강조용)</summary>
    public SetupItemViewModel? PickingItem => _pickItem;

    /// <summary>
    /// 칸·사진 클릭은 선택이 바뀌는 즉시 지정한다(자동).
    /// 여러 개 선택(그림 여러 개, 행 범위)은 다 고른 뒤 [확인]을 눌러야 한다.
    /// </summary>
    public bool PickNeedsConfirm => _pickKind is PickKind.Shapes or PickKind.Rows or PickKind.Lines;

    /// <summary>사진 옆 측정값 칸처럼 없을 수도 있는 단계는 [건너뛰기]</summary>
    public bool PickCanSkip => _pickKind == PickKind.ValueCell;

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
        SkipPickCommand    = new RelayCommand(_ => NextPickStep());
        CancelPickCommand  = new RelayCommand(_ => CancelPick());
        ScanPlaceholdersCommand = new RelayCommand(_ => ScanPlaceholders(showMessage: true));

        Items.CollectionChanged += (_, _) => OnPropertyChanged(nameof(HasNoItems));
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
            var item = new SetupItemViewModel(this, MakeId(label), label, GuessTypeForLabel(label)) { Placeholder = ph };
            Items.Add(item);
            added++;
        }

        if (showMessage || found.Count == 0)
            MessageBox.Show(found.Count == 0
                ? "서식에서 {{ }} 자리를 찾지 못했습니다.\n\n한글에서 채울 자리에 {{날짜}}, {{날씨}} 처럼 적어 두고 hwpx 로 저장한 뒤 다시 불러오세요.\n(자리 글자는 한 번에 입력하고, 중간에 글꼴을 바꾸지 마세요)"
                : $"자리 {found.Count}개를 찾았고, 새 항목 {added}개를 추가했습니다.\n종류가 다르면 항목을 지우고 [+ 항목 추가]로 다시 만드세요.",
                "한글 서식", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    /// <summary>자리 이름으로 종류 추측 (날짜·날씨·작업사항). 나머지는 매일 입력.</summary>
    private static string GuessTypeForLabel(string label) =>
        label.Contains("날짜") || label.Contains("일자") ? FieldTypes.DateSimple
        : label.Contains("날씨") ? FieldTypes.Text
        : FieldTypes.Input;

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
        Items.Remove(item);
    }

    public void RevealItem(SetupItemViewModel item)
    {
        if (IsPicking || Bridge == null) return;
        Bridge.Reveal(item.Type == FieldTypes.Photo ? item.Cell
                      : item.Type == FieldTypes.Multiline ? item.Cells.FirstOrDefault()
                      : item.Cell,
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
            PickKind.Column    => $"「{item.Label}」 {SetupItemViewModel.ColTitle(columnKey ?? "")} 열의 아무 칸이나 클릭하세요.",
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

            case PickKind.Picture when sel.IsShapes && sel.ShapeNames.Count == 1:
                item.Cell = sel.ActiveCell;
                NextPickStep();
                break;
        }
    }

    /// <summary>[확인]: 여러 개를 고르는 단계(그림 여러 개, 행 범위)를 현재 선택으로 마친다.</summary>
    private void ConfirmPick()
    {
        var item = _pickItem;
        var sel  = Bridge?.GetSelection();
        if (item == null || sel == null) return;

        switch (_pickKind)
        {
            case PickKind.Shapes:
                if (!sel.IsShapes) { Warn("그림·글상자를 먼저 클릭해 선택하세요."); return; }
                item.SetShapes(sel.ShapeNames, sel.ShapeIndexes);
                break;

            case PickKind.Rows:
                if (sel.IsShapes) { Warn("표의 칸들을 드래그해 선택하세요."); return; }
                item.SetRows(sel.Row, sel.Row + sel.RowCount - 1);
                break;

            case PickKind.Lines:
                if (sel.IsShapes) { Warn("칸들을 드래그해 선택하세요."); return; }
                string col = new(sel.ActiveCell.TakeWhile(char.IsLetter).ToArray());
                item.SetLines(Enumerable.Range(sel.Row, sel.RowCount).Select(r => $"{col}{r}").ToList());
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

        // 표: 행 범위 → 아직 안 정한 열을 순서대로
        if (_pickKind is PickKind.Rows or PickKind.Column)
        {
            string? nextCol = item.ColKeys.FirstOrDefault(k => item.ColOf(k) == null);
            if (nextCol != null) { StartPick(item, PickKind.Column, nextCol); return; }
        }

        CancelPick();

        if (_pickKind == PickKind.Shapes && item.Presets.Count == 0)
            MessageBox.Show("이제 Excel 에서 그림을 원하는 자리로 끌어다 놓고,\n항목 카드의 [지금 위치 저장]을 누르세요.\n위치를 여러 개 저장해 두면 매일 그중에서 고를 수 있습니다.",
                            "위치 저장 안내", MessageBoxButton.OK, MessageBoxImage.Information);
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
            var doc = new DocumentConfig
            {
                Id                = id,
                Name              = DocumentName.Trim(),
                TemplateFile      = _filePath,
                OutputPattern     = OutputPattern.Trim(),
                SheetName         = sheets.FirstOrDefault() ?? "",
                SheetNames        = sheets.Count > 1 ? sheets : [],
                Frequency         = SelectedFrequency.Value,
                FrequencyDays     = IsCustomFrequency ? FrequencyDays : 0,
                IncludesSujangjae = IncludesSujangjae,
                PhotoSheet        = PhotoSheet,
                PhotoSlots        = PhotoSlotRows.Select(r => r.ToConfig()).ToList(),
                Fields            = Items.Select(i => i.ToConfig()).ToList(),
                // 이 화면에서 다루지 않는 값은 기존 설정을 그대로 유지
                SaveMode          = _editing?.SaveMode ?? "create",
                PhotoFolder       = _editing?.PhotoFolder ?? "",
            };

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
        new("2주마다",     "biweekly"),
        new("매월",        "monthly"),
        new("사용자 지정", "custom"),
    ];
}
