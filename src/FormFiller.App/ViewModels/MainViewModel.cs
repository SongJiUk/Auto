using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Core.Models;

namespace FormFiller.App.ViewModels;

/// <summary>
/// MainWindow 의 DataContext.
/// DataContext: Unity의 MonoBehaviour 가 GameObject 에 붙는 것처럼,
/// ViewModel 이 화면(View)에 연결되어 데이터를 공급한다.
/// XAML 에서 {Binding PropertyName} 으로 여기 프로퍼티에 접근한다.
/// </summary>
public class MainViewModel : ViewModelBase
{
    // ── 기본 정보 ─────────────────────────────────────────────────
    private DateTime _targetDate = DateTime.Today;
    private string _weather = "맑음";

    public DateTime TargetDate
    {
        get => _targetDate;
        set
        {
            if (SetField(ref _targetDate, value))
                // 날짜가 바뀌면 "오늘 생성할 문서" 목록을 재계산한다
                LoadDocumentDueList();
        }
    }

    public string Weather
    {
        get => _weather;
        set => SetField(ref _weather, value);
    }

    // ── 날씨 선택 목록 ────────────────────────────────────────────
    public IReadOnlyList<string> WeatherOptions { get; } =
        ["맑음", "구름많음", "흐림", "비", "눈"];

    // ── 금일 작업사항 목록 ────────────────────────────────────────
    public ObservableCollection<WorkItemViewModel> TodayWorkItems { get; } = [];
    public ICommand AddTodayWorkCommand { get; }

    // ── 명일 작업사항 목록 ────────────────────────────────────────
    public ObservableCollection<WorkItemViewModel> TomorrowWorkItems { get; } = [];
    public ICommand AddTomorrowWorkCommand { get; }

    // ── 장비 투입현황 ─────────────────────────────────────────────
    public ObservableCollection<EquipmentEntryViewModel> Equipment { get; } =
    [
        new() { Name = "B/H 0.2", Spec = "5.8t"  },
        new() { Name = "B/H 380", Spec = "39.6t" },
        new() { Name = "B/H 480", Spec = "64.6t" },
        new() { Name = "B/H 1.0", Spec = "32.0t" },
    ];

    // ── 인원 투입현황 (공사팀) ────────────────────────────────────
    public ObservableCollection<PersonnelEntryViewModel> Personnel { get; } =
    [
        new() { Role = "관리자" },
        new() { Role = "벌목"   },
        new() { Role = "비계"   },
        new() { Role = "지하수" },
        new() { Role = "석면"   },
        new() { Role = "장비"   },
    ];

    // ── 수장재팀 별도 인원 ────────────────────────────────────────
    public ObservableCollection<PersonnelEntryViewModel> SujangjaePersonnel { get; } =
    [
        new() { Role = "수장재" },
    ];

    // ── 오늘 생성할 문서 목록 ─────────────────────────────────────
    public ObservableCollection<DocumentDueItem> DocumentDueList { get; } = [];

    // ── 사진 슬롯 ─────────────────────────────────────────────────
    /// <summary>
    /// 사진대지 등 이미지 슬롯이 있는 문서에서 자동으로 감지된 슬롯 목록.
    /// config.json의 각 문서에서 photoSheet 설정이 있으면 슬롯을 자동 생성한다.
    /// </summary>
    public ObservableCollection<PhotoSlotViewModel> PhotoSlots { get; } = [];

    /// <summary>사진 섹션 표시 여부 (슬롯이 1개 이상 있을 때만)</summary>
    public bool HasPhotoSlots => PhotoSlots.Count > 0;

    // ── 합계 (읽기 전용, 화면 하단에 표시) ──────────────────────
    public int TotalTodayEquipment  => Equipment.Sum(e => e.TodayCount);
    public int TotalTodayPersonnel  => Personnel.Sum(p => p.TodayCount);

    /// <summary>TBM 합계 = 공사인원 + 수장재인원</summary>
    public int TotalTodayPersonnelWithSujangjae =>
        Personnel.Sum(p => p.TodayCount) + SujangjaePersonnel.Sum(p => p.TodayCount);

    // ── 커맨드 ────────────────────────────────────────────────────
    public ICommand GenerateCommand { get; }

    public MainViewModel()
    {
        GenerateCommand        = new RelayCommand(_ => OnGenerate());
        AddTodayWorkCommand    = new RelayCommand(_ => AddWorkItem(TodayWorkItems));
        AddTomorrowWorkCommand = new RelayCommand(_ => AddWorkItem(TomorrowWorkItems));

        // 장비·인원 항목이 바뀔 때 합계도 갱신
        foreach (var e in Equipment)
            e.PropertyChanged += (_, _) => OnPropertyChanged(nameof(TotalTodayEquipment));
        foreach (var p in Personnel)
            p.PropertyChanged += (_, _) =>
            {
                OnPropertyChanged(nameof(TotalTodayPersonnel));
                OnPropertyChanged(nameof(TotalTodayPersonnelWithSujangjae));
            };

        // 수장재 인원 변경 시 TBM 합계 갱신
        foreach (var p in SujangjaePersonnel)
            p.PropertyChanged += (_, _) => OnPropertyChanged(nameof(TotalTodayPersonnelWithSujangjae));

        // 어제 파일에서 작업사항 불러오기 (없으면 기본 빈 항목 1개)
        LoadYesterdayWorkItems();

        // 어제 파일에서 장비 누계 / 인원 총누계 채우기
        LoadYesterdayData();

        // 문서 생성 주기 기반 오늘 목록 초기화
        LoadDocumentDueList();

        // 사진 슬롯 감지
        LoadPhotoSlots();
    }

    // ── 문서 목록 로드 ────────────────────────────────────────────

    /// <summary>
    /// config.json 의 문서 목록을 읽어 DocumentDueList 를 갱신한다.
    /// TargetDate 변경 시에도 호출되어 주기 판단을 재계산한다.
    /// </summary>
    private void LoadDocumentDueList()
    {
        DocumentDueList.Clear();
        var    config  = AppConfigLoader.Load();
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        foreach (var doc in config.Documents)
            DocumentDueList.Add(new DocumentDueItem(doc, TargetDate.Date, config.OutputFolder, baseDir));
    }

    // ── 파일 생성 ────────────────────────────────────────────────

    private void OnGenerate()
    {
        try
        {
            // ── 1. 설정 로드 ─────────────────────────────────────────
            var    config  = AppConfigLoader.Load();
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            // ── 2. 체크된 문서만 추출 ────────────────────────────────
            var checkedDocs = DocumentDueList.Where(d => d.IsChecked).ToList();
            if (checkedDocs.Count == 0)
            {
                MessageBox.Show("생성할 문서를 선택하세요.", "알림",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // ── 3. DailyInput 조립 ───────────────────────────────────
            var input = BuildDailyInput();

            // ── 4. 어제 파일에서 전일 누계 읽기 ─────────────────────
            var dailyDoc = config.Documents.FirstOrDefault(d => d.Id == "daily_report");
            if (dailyDoc != null)
            {
                DateTime yesterday    = TargetDate.AddDays(-1);
                string   prevFileName = GenericExcelWriter.ApplyDatePattern(
                                            dailyDoc.OutputPattern, yesterday);
                string   prevPath     = Path.Combine(baseDir, config.OutputFolder, prevFileName);

                GenericExcelWriter.LoadPreviousDay(prevPath, input, dailyDoc);
            }

            // ── 5. 체크된 문서만 생성 ───────────────────────────────
            var results = new List<string>();
            foreach (var dueItem in checkedDocs)
            {
                string outputPath = GenericExcelWriter.Generate(
                    input, dueItem.Config, baseDir, config.OutputFolder);
                results.Add($"{dueItem.Name}: {Path.GetFileName(outputPath)}");
            }

            // ── 6. 사진 삽입 ────────────────────────────────────────
            InsertPhotos(checkedDocs, config, baseDir, results);

            // ── 7. 완료 메시지 ───────────────────────────────────────
            string summary = string.Join("\n", results.Prepend("파일 생성 완료!\n"));
            MessageBox.Show(summary, "완료",
                            MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"파일 생성 중 오류가 발생했습니다.\n\n{ex.Message}",
                "오류",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// ViewModel 의 ObservableCollection 을 Core 모델(DailyInput)로 변환한다.
    /// SujangjaePersonnel 도 함께 포함시킨다.
    /// </summary>
    private DailyInput BuildDailyInput() => new()
    {
        Date         = TargetDate,
        Weather      = Weather,
        TodayWork    = string.Join("\n", TodayWorkItems.Select(i => i.Text)),
        TomorrowWork = string.Join("\n", TomorrowWorkItems.Select(i => i.Text)),
        Equipment    = Equipment.Select(e => new EquipmentEntry
        {
            Name         = e.Name,
            Spec         = e.Spec,
            PrevDayCount = e.PrevDayCount,
            TodayCount   = e.TodayCount,
        }).ToList(),
        Personnel    = Personnel.Select(p => new PersonnelEntry
        {
            Role                = p.Role,
            PrevCumulativeCount = p.PrevCumulativeCount,
            TodayCount          = p.TodayCount,
        }).ToList(),
        SujangjaePersonnel = SujangjaePersonnel.Select(p => new PersonnelEntry
        {
            Role                = p.Role,
            PrevCumulativeCount = p.PrevCumulativeCount,
            TodayCount          = p.TodayCount,
        }).ToList(),
    };

    // ── 작업사항 목록 초기화 ──────────────────────────────────────

    /// <summary>
    /// 어제 output 파일을 찾아 작업사항 셀 값을 목록에 채운다.
    /// 파일이 없거나 읽기에 실패하면 기본 빈 항목("- ")을 하나씩 추가한다.
    /// </summary>
    private void LoadYesterdayWorkItems()
    {
        var config   = AppConfigLoader.Load();
        var dailyDoc = config.Documents.FirstOrDefault(d => d.Id == "daily_report");
        if (dailyDoc == null)
        {
            InitDefaultItems();
            return;
        }

        DateTime yesterday = DateTime.Today.AddDays(-1);
        string   fileName  = GenericExcelWriter.ApplyDatePattern(dailyDoc.OutputPattern, yesterday);
        string   filePath  = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, config.OutputFolder, fileName);

        if (!File.Exists(filePath))
        {
            InitDefaultItems();
            return;
        }

        try
        {
            var todayField    = dailyDoc.Fields.FirstOrDefault(f => f.Id == "today_work");
            var tomorrowField = dailyDoc.Fields.FirstOrDefault(f => f.Id == "tomorrow_work");

            using var doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(
                filePath, isEditable: false);
            var wbPart = doc.WorkbookPart!;
            var wsPart = CellHelper.FindWorksheetPart(wbPart, dailyDoc.SheetName);
            if (wsPart == null)
            {
                InitDefaultItems();
                return;
            }

            LoadItemsFromCells(TodayWorkItems,    wsPart, todayField?.Cells);
            LoadItemsFromCells(TomorrowWorkItems, wsPart, tomorrowField?.Cells);
        }
        catch
        {
            InitDefaultItems();
        }
    }

    /// <summary>
    /// 지정한 셀 범위의 값을 읽어 items 컬렉션에 추가한다.
    /// 값이 하나도 없으면 빈 항목("- ")을 하나 넣는다.
    /// </summary>
    private void LoadItemsFromCells(
        ObservableCollection<WorkItemViewModel> items,
        DocumentFormat.OpenXml.Packaging.WorksheetPart wsPart,
        List<string>? cells)
    {
        items.Clear();
        if (cells == null)
        {
            items.Add(CreateWorkItem("- ", items));
            return;
        }

        bool hasAny = false;
        foreach (var cellAddr in cells)
        {
            string? val = CellHelper.GetCellValue(wsPart, cellAddr);
            if (!string.IsNullOrWhiteSpace(val))
            {
                items.Add(CreateWorkItem(val, items));
                hasAny = true;
            }
        }

        if (!hasAny)
            items.Add(CreateWorkItem("- ", items));
    }

    /// <summary>금일·명일 작업사항 모두 빈 항목 하나로 초기화한다.</summary>
    private void InitDefaultItems()
    {
        TodayWorkItems.Clear();
        TodayWorkItems.Add(CreateWorkItem("- ", TodayWorkItems));

        TomorrowWorkItems.Clear();
        TomorrowWorkItems.Add(CreateWorkItem("- ", TomorrowWorkItems));
    }

    /// <summary>WorkItemViewModel 을 생성하면서 삭제 콜백을 올바른 컬렉션에 연결한다.</summary>
    private static WorkItemViewModel CreateWorkItem(
        string text,
        ObservableCollection<WorkItemViewModel> owner)
        => new(text, item => owner.Remove(item));

    /// <summary>지정한 컬렉션에 빈 항목 하나를 추가하는 커맨드 핸들러.</summary>
    private static void AddWorkItem(ObservableCollection<WorkItemViewModel> items)
        => items.Add(new WorkItemViewModel("- ", item => items.Remove(item)));

    // ── 어제 누계 로드 ────────────────────────────────────────────

    /// <summary>
    /// 어제 output 파일을 열어 장비 누계(J열)와 인원 총누계(R열)를 읽은 뒤
    /// Equipment[i].PrevDayCount 와 Personnel[i].PrevCumulativeCount 에 채운다.
    /// 파일이 없거나 읽기에 실패하면 조용히 종료 (기존 0 값 유지).
    /// </summary>
    private void LoadYesterdayData()
    {
        var config   = AppConfigLoader.Load();
        var dailyDoc = config.Documents.FirstOrDefault(d => d.Id == "daily_report");
        if (dailyDoc == null) return;

        DateTime yesterday = DateTime.Today.AddDays(-1);
        string   fileName  = GenericExcelWriter.ApplyDatePattern(dailyDoc.OutputPattern, yesterday);
        string   filePath  = Path.Combine(
                                 AppDomain.CurrentDomain.BaseDirectory,
                                 config.OutputFolder,
                                 fileName);

        if (!File.Exists(filePath)) return;

        try
        {
            using var doc    = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(
                                   filePath, isEditable: false);
            var       wbPart = doc.WorkbookPart!;
            var       wsPart = CellHelper.FindWorksheetPart(wbPart, dailyDoc.SheetName);
            if (wsPart == null) return;

            var eqField = dailyDoc.Fields.FirstOrDefault(f => f.Type == FormFiller.Core.Config.FieldTypes.EquipmentTable);
            var pnField = dailyDoc.Fields.FirstOrDefault(f => f.Type == FormFiller.Core.Config.FieldTypes.PersonnelTable);

            // ── 장비 누계 읽기 (todayCol 에서 두 열 오른쪽 = 누계 열) ──
            if (eqField != null && !string.IsNullOrEmpty(eqField.NameCol)
                                && !string.IsNullOrEmpty(eqField.TodayCol))
            {
                string cumCol = NextNextColumn(eqField.TodayCol);
                for (int row = eqField.StartRow; row <= eqField.EndRow; row++)
                {
                    string? name = CellHelper.GetCellValue(wsPart, $"{eqField.NameCol}{row}");
                    if (string.IsNullOrWhiteSpace(name)) continue;

                    var vm = Equipment.FirstOrDefault(e => Norm(e.Name) == Norm(name));
                    if (vm == null) continue;

                    vm.PrevDayCount = CellHelper.GetCellInt(wsPart, $"{cumCol}{row}") ?? 0;
                }
            }

            // ── 인원 총누계 읽기 (todayCol 에서 두 열 오른쪽 = 총누계 열) ──
            if (pnField != null && !string.IsNullOrEmpty(pnField.RoleCol)
                                && !string.IsNullOrEmpty(pnField.TodayCol))
            {
                string cumCol = NextNextColumn(pnField.TodayCol);
                for (int row = pnField.StartRow; row <= pnField.EndRow; row++)
                {
                    string? role = CellHelper.GetCellValue(wsPart, $"{pnField.RoleCol}{row}");
                    if (string.IsNullOrWhiteSpace(role)) continue;

                    var vm = Personnel.FirstOrDefault(p => Norm(p.Role) == Norm(role));
                    if (vm == null) continue;

                    vm.PrevCumulativeCount = CellHelper.GetCellInt(wsPart, $"{cumCol}{row}") ?? 0;
                }
            }
        }
        catch
        {
            // 읽기 오류 시 조용히 무시 — 0 값 유지
        }
    }

    /// <summary>
    /// 열 문자를 두 칸 오른쪽으로 이동한다.
    /// 예) "H" → "J",  "P" → "R"
    /// </summary>
    private static string NextNextColumn(string col)
    {
        uint idx = CellHelper.ColLetterToIndex(col);
        return ColIndexToLetter(idx + 2);
    }

    /// <summary>
    /// 1-based 열 인덱스를 열 문자열로 변환한다.
    /// 예) 1→"A", 26→"Z", 27→"AA"
    /// </summary>
    private static string ColIndexToLetter(uint index)
    {
        string result = "";
        while (index > 0)
        {
            index--;
            result  = (char)('A' + index % 26) + result;
            index  /= 26;
        }
        return result;
    }

    /// <summary>이름/직종 비교를 위한 정규화: 앞뒤 공백 제거 + 소문자 변환</summary>
    private static string Norm(string s) => s.Trim().ToLowerInvariant();

    // ── 사진 슬롯 로드 ────────────────────────────────────────────

    /// <summary>
    /// config.json의 각 문서에서 PhotoSheet가 설정된 경우
    /// 해당 시트의 이미지 슬롯을 자동 감지해 PhotoSlots를 채운다.
    /// </summary>
    private void LoadPhotoSlots()
    {
        PhotoSlots.Clear();

        var    config  = AppConfigLoader.Load();
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        foreach (var doc in config.Documents)
        {
            if (string.IsNullOrWhiteSpace(doc.PhotoSheet)) continue;

            string templatePath = Path.IsPathRooted(doc.TemplateFile)
                ? doc.TemplateFile
                : Path.Combine(baseDir, doc.TemplateFile);

            if (!File.Exists(templatePath)) continue;

            var slotInfos = ExcelPhotoWriter.GetPhotoSlotInfos(templatePath, doc.PhotoSheet);

            foreach (var info in slotInfos)
            {
                // config에 저장된 슬롯 이름이 있으면 사용, 없으면 "사진 N"
                string? savedLabel = doc.PhotoSlots
                    .FirstOrDefault(s => s.Index == info.SlotIndex)?.Label;

                var slot = new PhotoSlotViewModel(
                    doc.Id, doc.PhotoSheet, info.SlotIndex,
                    info.WidthPx, info.HeightPx, savedLabel);
                PhotoSlots.Add(slot);
            }
        }

        OnPropertyChanged(nameof(HasPhotoSlots));
    }

    // ── 사진 삽입 ────────────────────────────────────────────────

    /// <summary>
    /// 선택된 사진 슬롯을 처리해 각 output 파일에 삽입한다.
    ///
    /// [흐름]
    /// 1. PhotoSlots에서 SelectedImagePath가 있는 슬롯만 추출
    /// 2. 해당 문서의 output 파일 경로 조회
    /// 3. PhotoService.ResizeToFill로 가득채우기 리사이즈
    /// 4. ExcelPhotoWriter.ReplacePhoto로 xlsx에 삽입
    /// </summary>
    private void InsertPhotos(
        List<DocumentDueItem> checkedDocs,
        AppConfig config,
        string baseDir,
        List<string> results)
    {
        var filledSlots = PhotoSlots
            .Where(s => s.HasPhoto)
            .ToList();

        if (filledSlots.Count == 0) return;

        foreach (var slot in filledSlots)
        {
            // 이 슬롯이 속한 문서의 output 경로 찾기
            var docConfig = config.Documents.FirstOrDefault(d => d.Id == slot.DocumentId);
            if (docConfig == null) continue;

            // 체크된 문서에 포함되어 있어야 함
            var dueItem = checkedDocs.FirstOrDefault(d => d.Config.Id == slot.DocumentId);
            if (dueItem == null) continue;

            string outputFileName = GenericExcelWriter.ApplyDatePattern(
                docConfig.OutputPattern, TargetDate);
            string outputPath = Path.Combine(baseDir, config.OutputFolder, outputFileName);

            if (!File.Exists(outputPath)) continue;

            try
            {
                // 가득채우기 리사이즈 → JPEG 바이너리
                byte[] imageBytes = PhotoService.ResizeToFill(
                    slot.SelectedImagePath!, slot.TargetWidthPx, slot.TargetHeightPx);

                // xlsx 이미지 슬롯 교체
                ExcelPhotoWriter.ReplacePhoto(
                    outputPath, slot.SheetName, slot.SlotIndex, imageBytes);

                results.Add($"  └ 사진 삽입: {slot.Label} → {slot.SheetName}");
            }
            catch (Exception ex)
            {
                results.Add($"  └ 사진 삽입 실패({slot.Label}): {ex.Message}");
            }
        }
    }
}
