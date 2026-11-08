using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using Microsoft.Win32;

namespace FormFiller.App.ViewModels.Settings;

/// <summary>
/// 문서 설정 창(DocumentSetupWindow) 의 ViewModel.
/// 사용자가 엑셀 파일을 선택하고, 시트를 고르고, 각 필드에 셀을 지정해서
/// 새로운 DocumentConfig 를 config.json 에 추가하는 역할을 담당한다.
/// </summary>
public class DocumentSetupViewModel : ViewModelBase
{
    // ════════════════════════════════════════════════════════════════
    // ── 파일 / 시트 ───────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    private string _filePath = "";

    /// <summary>
    /// 선택된 엑셀 파일 경로.
    /// 값이 바뀌면 시트 목록을 자동으로 갱신한다.
    /// </summary>
    public string FilePath
    {
        get => _filePath;
        set
        {
            if (SetField(ref _filePath, value))
                ReloadSheets();
        }
    }

    /// <summary>파일에서 읽어 온 시트 이름 목록 (ComboBox 에 바인딩)</summary>
    public ObservableCollection<string> SheetNames { get; } = [];

    private string _selectedSheet = "";

    /// <summary>
    /// 선택된 시트.
    /// 바뀌면 엑셀 미리보기 데이터를 갱신한다.
    /// </summary>
    public string SelectedSheet
    {
        get => _selectedSheet;
        set
        {
            if (SetField(ref _selectedSheet, value))
                ReloadPreviewData();
        }
    }

    // ════════════════════════════════════════════════════════════════
    // ── 문서 기본 정보 ────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    private string _documentName = "";

    /// <summary>문서 이름 (예: "공사일보_출력일보")</summary>
    public string DocumentName
    {
        get => _documentName;
        set => SetField(ref _documentName, value);
    }

    private string _outputPattern = "파일명_{yy}.{MM}.{dd}.xlsx";

    /// <summary>
    /// 출력 파일명 패턴.
    /// {yy}=연도 2자리, {MM}=월, {dd}=일 로 치환된다.
    /// </summary>
    public string OutputPattern
    {
        get => _outputPattern;
        set => SetField(ref _outputPattern, value);
    }

    // ── 생성 주기 ──────────────────────────────────────────────────

    /// <summary>주기 선택 항목 목록 (ComboBox 에 바인딩)</summary>
    public IReadOnlyList<FrequencyOption> FrequencyOptions { get; } = FrequencyOption.All;

    private FrequencyOption _selectedFrequency = FrequencyOption.All[0];

    /// <summary>선택된 주기 옵션</summary>
    public FrequencyOption SelectedFrequency
    {
        get => _selectedFrequency;
        set
        {
            if (SetField(ref _selectedFrequency, value))
                OnPropertyChanged(nameof(IsCustomFrequency));
        }
    }

    /// <summary>사용자 지정 주기(일수 입력 필드 표시 여부)</summary>
    public bool IsCustomFrequency => _selectedFrequency?.Value == "custom";

    private int _frequencyDays = 7;

    /// <summary>사용자 지정 주기 일수 (IsCustomFrequency=true 일 때 사용)</summary>
    public int FrequencyDays
    {
        get => _frequencyDays;
        set => SetField(ref _frequencyDays, value);
    }

    // ── 수장재팀 포함 여부 ─────────────────────────────────────────

    private bool _includesSujangjae;

    /// <summary>이 문서의 인원에 수장재팀을 포함할지 여부</summary>
    public bool IncludesSujangjae
    {
        get => _includesSujangjae;
        set => SetField(ref _includesSujangjae, value);
    }

    // ── 사진 시트 ────────────────────────────────────────────────────

    private const string NoPhotoOption = "없음 (사진 기능 미사용)";

    /// <summary>
    /// 사진 시트 선택 목록.
    /// "없음 (사진 기능 미사용)" + SheetNames 로 구성되며,
    /// 파일/시트 목록이 바뀔 때마다 갱신된다.
    /// </summary>
    public ObservableCollection<string> PhotoSheetOptions { get; } = [NoPhotoOption];

    private string _selectedPhotoSheet = NoPhotoOption;

    /// <summary>선택된 사진 시트. "없음…"을 고르면 PhotoSheet = "" 로 저장된다.</summary>
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

    /// <summary>저장용 실제 값. "없음…" 선택이면 빈 문자열.</summary>
    public string PhotoSheet =>
        _selectedPhotoSheet == NoPhotoOption ? "" : _selectedPhotoSheet ?? "";

    // ── 사진 슬롯 이름 목록 ──────────────────────────────────────────

    /// <summary>선택된 사진 시트에서 감지된 슬롯 목록. 각 슬롯에 이름을 붙일 수 있다.</summary>
    public ObservableCollection<PhotoSlotLabelRow> PhotoSlotRows { get; } = [];

    // ════════════════════════════════════════════════════════════════
    // ── 엑셀 미리보기 데이터 ──────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    private ExcelPreviewData? _previewFull;

    /// <summary>
    /// 엑셀 미리보기 전체 데이터 (셀 값·배경색·정렬·볼드·테두리·열 너비·행 높이·병합 셀).
    /// View(code-behind) 가 이 객체 하나로 Grid 를 완성한다.
    /// null 이면 미리보기 없음.
    /// </summary>
    public ExcelPreviewData? PreviewFull
    {
        get => _previewFull;
        private set
        {
            _previewFull = value;
            OnPropertyChanged();
        }
    }

    // ════════════════════════════════════════════════════════════════
    // ── 선택된 셀 정보 ────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    private string _selectedCellAddress = "";

    /// <summary>현재 클릭된 셀 주소 (예: "C6")</summary>
    public string SelectedCellAddress
    {
        get => _selectedCellAddress;
        set => SetField(ref _selectedCellAddress, value);
    }

    private string _selectedCellValue = "";

    /// <summary>현재 클릭된 셀의 값</summary>
    public string SelectedCellValue
    {
        get => _selectedCellValue;
        set => SetField(ref _selectedCellValue, value);
    }

    // ════════════════════════════════════════════════════════════════
    // ── 피커 모드 ─────────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    private FieldAssignRow? _activePickerField;

    /// <summary>
    /// 현재 셀 선택을 기다리는 필드.
    /// null 이면 피커 모드 비활성 상태.
    /// View 에서 이 값을 보고 "셀을 클릭하세요" 안내를 표시한다.
    /// </summary>
    public FieldAssignRow? ActivePickerField
    {
        get => _activePickerField;
        set
        {
            if (SetField(ref _activePickerField, value))
                OnPropertyChanged(nameof(IsPickerActive));
        }
    }

    /// <summary>피커 모드 활성 여부 (안내 패널 표시용)</summary>
    public bool IsPickerActive => _activePickerField != null;

    // ════════════════════════════════════════════════════════════════
    // ── 이벤트 ────────────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    /// <summary>저장 성공 시 View 에 창 닫기를 요청하는 이벤트</summary>
    public event EventHandler? SaveSucceeded;

    // ════════════════════════════════════════════════════════════════
    // ── 필드 목록 ─────────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    /// <summary>
    /// 사용자가 셀을 지정할 필드 목록.
    /// 빈 컬렉션으로 시작하고, 사용자가 직접 추가한다.
    /// </summary>
    public ObservableCollection<FieldAssignRow> Fields { get; } = [];

    // ════════════════════════════════════════════════════════════════
    // ── 새 필드 추가 패널용 임시 프로퍼티 ────────────────────────
    // ════════════════════════════════════════════════════════════════

    private bool _showAddFieldPanel;

    /// <summary>새 필드 추가 패널 표시 여부</summary>
    public bool ShowAddFieldPanel
    {
        get => _showAddFieldPanel;
        set => SetField(ref _showAddFieldPanel, value);
    }

    private string _newFieldLabel = "";

    /// <summary>추가할 필드의 표시 이름</summary>
    public string NewFieldLabel
    {
        get => _newFieldLabel;
        set => SetField(ref _newFieldLabel, value);
    }

    private FieldTypeOption? _newFieldType;

    /// <summary>추가할 필드의 타입 선택 항목</summary>
    public FieldTypeOption? NewFieldType
    {
        get => _newFieldType;
        set => SetField(ref _newFieldType, value);
    }

    /// <summary>콤보박스에 표시할 전체 타입 목록</summary>
    public IReadOnlyList<FieldTypeOption> FieldTypeOptions => FieldTypeOption.All;

    // ════════════════════════════════════════════════════════════════
    // ── 커맨드 ────────────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    /// <summary>파일 선택 대화상자를 열어 엑셀 파일을 선택한다</summary>
    public ICommand BrowseFileCommand { get; }

    /// <summary>설정을 config.json 에 새 DocumentConfig 로 추가·저장한다</summary>
    public ICommand SaveCommand { get; }

    /// <summary>새 필드 추가 패널을 열어 입력 대기 상태로 만든다</summary>
    public ICommand AddFieldCommand { get; }

    /// <summary>입력된 정보로 새 필드를 목록에 추가하고 패널을 닫는다</summary>
    public ICommand ConfirmAddFieldCommand { get; }

    /// <summary>새 필드 추가를 취소하고 패널을 닫는다</summary>
    public ICommand CancelAddFieldCommand { get; }

    // ════════════════════════════════════════════════════════════════
    // ── 생성자 ────────────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    public DocumentSetupViewModel()
    {
        // ── 커맨드 연결 ──
        BrowseFileCommand = new RelayCommand(_ => OnBrowseFile());
        SaveCommand       = new RelayCommand(_ => OnSave());

        // 필드 추가 패널 열기: 입력 초기화 후 패널 표시
        AddFieldCommand = new RelayCommand(_ =>
        {
            NewFieldLabel = "";
            NewFieldType  = FieldTypeOption.All[0];
            ShowAddFieldPanel = true;
        });

        // 추가 확인: 이름·타입 검증 후 목록에 삽입
        ConfirmAddFieldCommand = new RelayCommand(_ =>
        {
            if (string.IsNullOrWhiteSpace(NewFieldLabel) || NewFieldType == null) return;

            // 필드 ID: 이름에서 공백·괄호를 제거해 소문자로 만든다
            string fieldId = NewFieldLabel
                .ToLowerInvariant()
                .Replace(" ", "_")
                .Replace("(", "")
                .Replace(")", "");

            // 삭제 콜백: 해당 항목을 Fields 컬렉션에서 제거
            var newField = new FieldAssignRow(
                fieldId,
                NewFieldLabel,
                NewFieldType.TypeId,
                StartPick,
                item => Fields.Remove(item));

            Fields.Add(newField);
            ShowAddFieldPanel = false;
        });

        // 추가 취소: 패널만 닫는다
        CancelAddFieldCommand = new RelayCommand(_ => ShowAddFieldPanel = false);
    }

    /// <summary>
    /// FieldAssignRow 의 StartPickCommand 에서 호출하는 피커 활성화 콜백.
    /// 테이블 타입이 아닌 경우에만 ActivePickerField 를 설정한다.
    /// </summary>
    private void StartPick(FieldAssignRow field)
    {
        if (!field.IsTableType)
            ActivePickerField = field;
    }

    // ════════════════════════════════════════════════════════════════
    // ── public 메서드: View 에서 호출 ────────────────────────────
    // ════════════════════════════════════════════════════════════════

    /// <summary>
    /// 엑셀 미리보기에서 셀을 클릭했을 때 View(code-behind) 에서 호출한다.
    ///
    /// [처리 흐름]
    /// 1. 선택된 셀 주소/값을 저장 (하단 정보 표시용)
    /// 2. 피커 모드가 활성이고 대상 필드가 테이블 타입이 아니면 → 셀 주소를 필드에 저장하고 피커 모드 종료
    /// </summary>
    public void OnCellClicked(string cellAddress, string cellValue)
    {
        // 선택된 셀 정보 갱신 (항상)
        SelectedCellAddress = cellAddress;
        SelectedCellValue   = cellValue;

        // 피커 모드 활성 + 단순 셀 타입이면 자동으로 필드에 주소 입력
        if (ActivePickerField != null && !ActivePickerField.IsTableType)
        {
            ActivePickerField.AssignedCell = cellAddress;
            ActivePickerField = null; // 피커 모드 종료
        }
    }

    // ════════════════════════════════════════════════════════════════
    // ── private 메서드 ────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    /// <summary>OpenFileDialog 로 .xlsx 파일을 선택한다.</summary>
    private void OnBrowseFile()
    {
        var dialog = new OpenFileDialog
        {
            Title  = "엑셀 서식 파일 선택",
            Filter = "Excel 파일 (*.xlsx)|*.xlsx",
        };

        if (dialog.ShowDialog() == true)
            FilePath = dialog.FileName;
    }

    /// <summary>
    /// 파일 경로가 바뀌었을 때 시트 목록을 다시 읽는다.
    /// 첫 번째 시트를 자동으로 선택한다.
    /// </summary>
    private void ReloadSheets()
    {
        SheetNames.Clear();
        foreach (var s in ExcelAnalyzer.GetSheetNames(_filePath))
            SheetNames.Add(s);

        SelectedSheet = SheetNames.Count > 0 ? SheetNames[0] : "";

        // 사진 시트 선택 목록도 갱신
        PhotoSheetOptions.Clear();
        PhotoSheetOptions.Add(NoPhotoOption);
        foreach (var s in SheetNames)
            PhotoSheetOptions.Add(s);

        // 이전에 고른 값이 목록에 없으면 "없음"으로 초기화
        if (!PhotoSheetOptions.Contains(_selectedPhotoSheet))
            SelectedPhotoSheet = NoPhotoOption;

        ReloadPhotoSlotRows();
    }

    /// <summary>
    /// 선택된 사진 시트에서 이미지 슬롯을 감지해 PhotoSlotRows를 갱신한다.
    /// </summary>
    private void ReloadPhotoSlotRows()
    {
        PhotoSlotRows.Clear();

        if (string.IsNullOrEmpty(PhotoSheet)
            || string.IsNullOrWhiteSpace(_filePath)
            || !System.IO.File.Exists(_filePath)) return;

        try
        {
            var slotInfos = ExcelPhotoWriter.GetPhotoSlotInfos(_filePath, PhotoSheet);
            foreach (var slot in slotInfos)
            {
                byte[]? imgBytes = ExcelPhotoWriter.GetSlotImageBytes(_filePath, PhotoSheet, slot.SlotIndex);
                PhotoSlotRows.Add(new PhotoSlotLabelRow(slot.SlotIndex, slot.WidthPx, slot.HeightPx, "", imgBytes));
            }
        }
        catch { /* 파일 접근 실패 시 무시 */ }
    }

    /// <summary>
    /// 시트가 바뀌었을 때 엑셀 미리보기 데이터를 다시 읽는다.
    /// GetPreviewData 로 모든 스타일 정보를 한 번에 가져온다.
    /// 실패 시 셀 값만 읽는 간단한 방식으로 폴백한다.
    /// </summary>
    private void ReloadPreviewData()
    {
        if (string.IsNullOrWhiteSpace(_filePath) || string.IsNullOrWhiteSpace(_selectedSheet))
        {
            PreviewFull = null;
            return;
        }

        // 1차 시도: 스타일·병합·너비 포함 전체 데이터
        var full = ExcelAnalyzer.GetPreviewData(_filePath, _selectedSheet);
        if (full != null)
        {
            PreviewFull = full;
            return;
        }

        // 2차 폴백: 셀 값만 읽어서 최소한 그리드는 보이게 한다
        System.Diagnostics.Debug.WriteLine(
            "[DocumentSetupVM] GetPreviewData 실패 → 셀 값 폴백");

        (int maxRow, int maxCol) = ExcelAnalyzer.GetUsedRange(_filePath, _selectedSheet);
        if (maxRow > 0 && maxCol > 0)
        {
            var values = ExcelAnalyzer.GetCellValues(_filePath, _selectedSheet, maxRow, maxCol);
            int rows   = values.GetLength(0);
            int cols   = values.GetLength(1);

            PreviewFull = new ExcelPreviewData
            {
                Rows         = rows,
                Cols         = cols,
                Values       = values,
                BgColors     = new uint?[rows, cols],
                HAlignments  = new ExcelHAlign[rows, cols],
                IsBold       = new bool[rows, cols],
                Borders      = new ExcelBorderInfo[rows, cols],
                ColWidthsPx  = Enumerable.Repeat(80.0, cols).ToArray(),
                RowHeightsPx = Enumerable.Repeat(22.0, rows).ToArray(),
                Merges       = [],
            };
        }
        else
        {
            PreviewFull = null;
        }
    }

    /// <summary>
    /// 드래그 앤 드롭으로 파일을 받았을 때 View(code-behind) 에서 호출한다.
    /// files 배열 중 첫 번째 .xlsx 파일을 FilePath 로 설정한다.
    /// </summary>
    public void HandleDroppedFile(string[] files)
    {
        string? xlsx = files.FirstOrDefault(
            f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase));
        if (xlsx != null)
            FilePath = xlsx;
    }

    /// <summary>
    /// 현재 편집된 필드 설정으로 DocumentConfig 를 생성해 config.json 에 추가·저장한다.
    ///
    /// [저장 흐름]
    /// 1. AppConfigLoader.Load() 로 현재 config 로드
    /// 2. Fields 를 순회해 FieldConfig 목록 생성
    ///    - IsTableType → StartRow/EndRow/열 정보 포함
    ///    - Multiline   → AssignedCell 기준으로 9개 셀 주소 자동 생성
    ///    - 나머지       → Cell 주소만
    /// 3. 새 DocumentConfig 를 documents 에 추가
    /// 4. AppConfigLoader.Save() 로 저장
    /// </summary>
    private void OnSave()
    {
        // ── 기본값 유효성 검사 ──
        if (string.IsNullOrWhiteSpace(DocumentName))
        {
            MessageBox.Show("문서 이름을 입력해 주세요.", "입력 오류",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(FilePath))
        {
            MessageBox.Show("엑셀 파일을 선택해 주세요.", "입력 오류",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            var config = AppConfigLoader.Load();

            // ── DocumentConfig.Id: 이름에서 공백을 _로 치환해 간단한 식별자 생성 ──
            string docId = DocumentName
                .Replace(" ", "_")
                .Replace("/", "_")
                .ToLowerInvariant();

            // 중복 Id 방지: 이미 같은 Id 가 있으면 번호 추가
            string uniqueId = docId;
            int    suffix   = 2;
            while (config.Documents.Any(d => d.Id == uniqueId))
                uniqueId = $"{docId}_{suffix++}";

            // ── FieldConfig 목록 생성 ──
            var fieldConfigs = new List<FieldConfig>();

            foreach (var field in Fields)
            {
                if (field.IsTableType)
                {
                    // 테이블 타입: 행·열 정보를 포함한 FieldConfig
                    var fc = new FieldConfig
                    {
                        Id       = field.FieldId,
                        Label    = field.Label,
                        Type     = field.Type,
                        StartRow = field.StartRow,
                        EndRow   = field.EndRow,
                    };

                    if (field.Type == FieldTypes.EquipmentTable)
                    {
                        fc.NameCol    = field.NameCol;
                        fc.PrevDayCol = field.PrevDayCol;
                        fc.TodayCol   = field.TodayCol;
                    }
                    else // PersonnelTable
                    {
                        fc.RoleCol           = field.RoleCol;
                        fc.PrevCumulativeCol = field.PrevCumulativeCol;
                        fc.TodayCol          = field.TodayCol;
                    }

                    fieldConfigs.Add(fc);
                }
                else if (field.Type == FieldTypes.Multiline)
                {
                    // Multiline: 시작 셀에서 9개 연속 셀 자동 생성 (예: A9 → A9~A17)
                    fieldConfigs.Add(new FieldConfig
                    {
                        Id    = field.FieldId,
                        Label = field.Label,
                        Type  = field.Type,
                        Cells = BuildCellRange(field.AssignedCell, 9),
                    });
                }
                else
                {
                    // 단일 셀 타입: Cell 주소만
                    fieldConfigs.Add(new FieldConfig
                    {
                        Id    = field.FieldId,
                        Label = field.Label,
                        Type  = field.Type,
                        Cell  = string.IsNullOrWhiteSpace(field.AssignedCell) ? null : field.AssignedCell,
                    });
                }
            }

            // ── 새 DocumentConfig 추가 ──
            var newDoc = new DocumentConfig
            {
                Id                = uniqueId,
                Name              = DocumentName,
                TemplateFile      = FilePath,
                OutputPattern     = OutputPattern,
                SheetName         = SelectedSheet,
                Frequency         = SelectedFrequency.Value,
                FrequencyDays     = IsCustomFrequency ? FrequencyDays : 0,
                IncludesSujangjae = IncludesSujangjae,
                PhotoSheet        = PhotoSheet,
                PhotoSlots        = PhotoSlotRows.Select(r => r.ToConfig()).ToList(),
                Fields            = fieldConfigs,
            };

            config.Documents.Add(newDoc);
            AppConfigLoader.Save(config);

            MessageBox.Show($"문서 '{DocumentName}' 이(가) 추가되었습니다.", "저장 완료",
                            MessageBoxButton.OK, MessageBoxImage.Information);

            SaveSucceeded?.Invoke(this, EventArgs.Empty);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"저장 중 오류가 발생했습니다.\n{ex.Message}", "저장 오류",
                            MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 시작 셀 주소에서 열 문자와 시작 행을 추출해 연속 셀 주소 목록을 생성한다.
    /// 예) BuildCellRange("A9", 9) → ["A9","A10",...,"A17"]
    /// </summary>
    private static List<string> BuildCellRange(string startCell, int count)
    {
        if (string.IsNullOrWhiteSpace(startCell)) return [];

        int i = 0;
        while (i < startCell.Length && char.IsLetter(startCell[i])) i++;

        string col = startCell[..i].ToUpperInvariant();
        if (!int.TryParse(startCell[i..], out int startRow)) return [];

        return Enumerable.Range(startRow, count)
                         .Select(r => $"{col}{r}")
                         .ToList();
    }
}

// ════════════════════════════════════════════════════════════════
// ── FrequencyOption: 주기 선택 항목 ──────────────────────────
// ════════════════════════════════════════════════════════════════

/// <summary>생성 주기 콤보박스 항목</summary>
public class FrequencyOption
{
    public string Label { get; }
    public string Value { get; }

    public FrequencyOption(string label, string value) { Label = label; Value = value; }

    public static readonly IReadOnlyList<FrequencyOption> All =
    [
        new("매일",         "daily"),
        new("2주마다",       "biweekly"),
        new("매월",         "monthly"),
        new("사용자 지정",    "custom"),
    ];
}

// ════════════════════════════════════════════════════════════════
// ── FieldAssignRow: 필드 목록의 각 항목 모델 ──────────────────
// ════════════════════════════════════════════════════════════════

/// <summary>
/// 문서 설정 화면 우측 패널에서 하나의 필드 행을 나타내는 모델.
/// 각 행은 필드 ID·레이블·타입·지정된 셀 주소로 구성된다.
/// 테이블 타입(장비표/인원표)은 추가로 행·열 범위 정보를 갖는다.
/// </summary>
public class FieldAssignRow : ViewModelBase
{
    // ── 불변 식별 정보 ─────────────────────────────────────────────

    /// <summary>필드 식별자 (예: "date", "equipment_table")</summary>
    public string FieldId { get; }

    /// <summary>사람이 읽을 필드 이름 (예: "날짜", "장비 투입현황")</summary>
    public string Label { get; }

    /// <summary>처리 타입 — FieldTypes 상수</summary>
    public string Type { get; }

    // ── 지정된 셀 주소 ─────────────────────────────────────────────

    private string _assignedCell = "";

    /// <summary>
    /// 사용자가 셀 클릭 또는 직접 입력으로 지정한 셀 주소 (예: "C6").
    /// 테이블 타입에서는 사용되지 않는다.
    /// </summary>
    public string AssignedCell
    {
        get => _assignedCell;
        set
        {
            if (SetField(ref _assignedCell, value))
            {
                OnPropertyChanged(nameof(IsAssigned));
                OnPropertyChanged(nameof(StatusText));
            }
        }
    }

    /// <summary>셀이 지정되었는지 여부 (UI 체크 표시용)</summary>
    public bool IsAssigned => !string.IsNullOrEmpty(_assignedCell);

    /// <summary>테이블 타입 여부 (피커 클릭 대신 "설정" 다이얼로그용)</summary>
    public bool IsTableType => Type is FieldTypes.EquipmentTable or FieldTypes.PersonnelTable;

    /// <summary>단순 셀 지정 타입 여부 (테이블 타입의 반대)</summary>
    public bool IsSimpleType => !IsTableType;

    /// <summary>
    /// 하단 상태 텍스트 — 셀이 지정됐으면 주소, 아니면 안내 문구.
    /// </summary>
    public string StatusText => IsAssigned ? $"지정됨: {AssignedCell}" : "미지정";

    // ── 테이블 타입 전용 설정 ──────────────────────────────────────

    private int _startRow;

    /// <summary>테이블 시작 행 (1-based)</summary>
    public int StartRow
    {
        get => _startRow;
        set => SetField(ref _startRow, value);
    }

    private int _endRow;

    /// <summary>테이블 끝 행 (합계 행 제외)</summary>
    public int EndRow
    {
        get => _endRow;
        set => SetField(ref _endRow, value);
    }

    private string _nameCol = "";

    /// <summary>장비명 열 (equipment_table 전용)</summary>
    public string NameCol
    {
        get => _nameCol;
        set => SetField(ref _nameCol, value);
    }

    private string _prevDayCol = "";

    /// <summary>전일(대) 열 (equipment_table 전용)</summary>
    public string PrevDayCol
    {
        get => _prevDayCol;
        set => SetField(ref _prevDayCol, value);
    }

    private string _todayCol = "";

    /// <summary>금일(대/명) 열 (공통)</summary>
    public string TodayCol
    {
        get => _todayCol;
        set => SetField(ref _todayCol, value);
    }

    private string _roleCol = "";

    /// <summary>직종 열 (personnel_table 전용)</summary>
    public string RoleCol
    {
        get => _roleCol;
        set => SetField(ref _roleCol, value);
    }

    private string _prevCumulativeCol = "";

    /// <summary>전일누계(명) 열 (personnel_table 전용)</summary>
    public string PrevCumulativeCol
    {
        get => _prevCumulativeCol;
        set => SetField(ref _prevCumulativeCol, value);
    }

    // ── 커맨드 ─────────────────────────────────────────────────────

    /// <summary>
    /// 피커 모드 시작 커맨드.
    /// 클릭 시 onStartPick 콜백을 통해 부모 ViewModel 에 알린다.
    /// </summary>
    public ICommand StartPickCommand { get; }

    /// <summary>
    /// 이 필드를 목록에서 제거하는 커맨드.
    /// onRemove 콜백이 없으면 아무 동작도 하지 않는다.
    /// </summary>
    public ICommand RemoveCommand { get; }

    // ── 생성자 ─────────────────────────────────────────────────────

    /// <summary>
    /// FieldAssignRow 생성자.
    /// </summary>
    /// <param name="fieldId">필드 ID</param>
    /// <param name="label">표시 이름</param>
    /// <param name="type">FieldTypes 상수</param>
    /// <param name="onStartPick">피커 모드 시작 요청 콜백</param>
    /// <param name="onRemove">이 행 삭제 요청 콜백 (optional)</param>
    public FieldAssignRow(string fieldId, string label, string type,
        Action<FieldAssignRow> onStartPick,
        Action<FieldAssignRow>? onRemove = null)
    {
        FieldId = fieldId;
        Label   = label;
        Type    = type;

        // 피커 모드 시작 커맨드
        StartPickCommand = new RelayCommand(_ => onStartPick(this));

        // 삭제 커맨드: 콜백이 있을 때만 동작
        RemoveCommand = new RelayCommand(_ => onRemove?.Invoke(this));
    }
}
