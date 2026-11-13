using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using FormFiller.App.Views;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Core.Models;

namespace FormFiller.App.ViewModels;

/// <summary>
/// MainWindow 의 DataContext.
/// </summary>
public class MainViewModel : ViewModelBase
{
    // ── 기본 정보 ─────────────────────────────────────────────────────
    private DateTime _targetDate = DateTime.Today;
    private string   _weather    = "맑음";

    public DateTime TargetDate
    {
        get => _targetDate;
        set
        {
            if (SetField(ref _targetDate, value))
                LoadDocumentDueList();
        }
    }

    public string Weather
    {
        get => _weather;
        set => SetField(ref _weather, value);
    }

    public IReadOnlyList<string> WeatherOptions { get; } =
        ["맑음", "구름많음", "흐림", "비", "눈"];

    // ── 금일 작업사항 (오전 / 오후) ──────────────────────────────────
    public ObservableCollection<WorkItemViewModel> TodayMorningItems   { get; } = [];
    public ObservableCollection<WorkItemViewModel> TodayAfternoonItems { get; } = [];

    // ── 명일 작업사항 ─────────────────────────────────────────────────
    public ObservableCollection<WorkItemViewModel> TomorrowWorkItems { get; } = [];

    // ── 장비 투입현황 ─────────────────────────────────────────────────
    public ObservableCollection<EquipmentEntryViewModel> Equipment { get; } = [];

    // ── 인원 투입현황 (공사팀) ────────────────────────────────────────
    public ObservableCollection<PersonnelEntryViewModel> Personnel { get; } = [];

    // ── 수장재팀 별도 인원 ────────────────────────────────────────────
    public ObservableCollection<PersonnelEntryViewModel> SujangjaePersonnel { get; } = [];

    // ── 오늘 생성할 문서 목록 ─────────────────────────────────────────
    public ObservableCollection<DocumentDueItem> DocumentDueList { get; } = [];

    // ── 사진 슬롯 ─────────────────────────────────────────────────────
    public ObservableCollection<PhotoSlotViewModel> PhotoSlots { get; } = [];
    public bool HasPhotoSlots => PhotoSlots.Count > 0;

    // 문서별 입력 (소음측정 사진·측정값·장비 위치 등, 시트마다 따로)
    public ObservableCollection<DocInputCardViewModel> DocInputCards { get; } = [];
    public bool HasDocInputs => DocInputCards.Count > 0;

    /// <summary>설정이 바뀐 뒤(문서 추가·편집) 메인 화면의 문서 목록·입력 칸을 다시 읽는다.</summary>
    public void ReloadDocuments()
    {
        LoadDocumentDueList();
        LoadPhotoSlots();
        LoadDocInputCards();
    }

    private void LoadDocInputCards()
    {
        DocInputCards.Clear();
        foreach (var doc in AppConfigLoader.Load().Documents
                     .Where(d => d.Fields.Any(f => FieldTypes.IsDailyInput(f.Type))))
            DocInputCards.Add(new DocInputCardViewModel(doc, () => TargetDate));
        OnPropertyChanged(nameof(HasDocInputs));
    }

    // ── 합계 (읽기 전용) ──────────────────────────────────────────────
    public int TotalTodayEquipment  => Equipment.Sum(e => e.TodayCount);
    public int TotalTodayPersonnel  => Personnel.Sum(p => p.TodayCount);
    public int TotalTodayPersonnelWithSujangjae =>
        Personnel.Sum(p => p.TodayCount) + SujangjaePersonnel.Sum(p => p.TodayCount);

    // ── 커맨드 ────────────────────────────────────────────────────────
    public ICommand GenerateCommand { get; }

    // 작업사항
    public ICommand AddMorningWorkCommand       { get; }
    public ICommand AddAfternoonWorkCommand     { get; }
    public ICommand AddTomorrowWorkCommand      { get; }
    public ICommand PickMorningWorkCommand      { get; }
    public ICommand PickAfternoonWorkCommand    { get; }
    public ICommand PickTomorrowWorkCommand     { get; }

    // 장비 추가
    public ICommand AddEquipmentCommand         { get; }

    // 인원 추가
    public ICommand AddPersonnelCommand         { get; }
    public ICommand AddSujangjaeCommand         { get; }

    // AI 도우미
    public ICommand OpenAiAssistantCommand      { get; }

    public MainViewModel()
    {
        GenerateCommand          = new RelayCommand(_ => OnGenerate());

        AddMorningWorkCommand    = new RelayCommand(_ => AddWorkItem(TodayMorningItems));
        AddAfternoonWorkCommand  = new RelayCommand(_ => AddWorkItem(TodayAfternoonItems));
        AddTomorrowWorkCommand   = new RelayCommand(_ => AddWorkItem(TomorrowWorkItems));

        PickMorningWorkCommand   = new RelayCommand(_ => OpenWorkPicker(TodayMorningItems));
        PickAfternoonWorkCommand = new RelayCommand(_ => OpenWorkPicker(TodayAfternoonItems));
        PickTomorrowWorkCommand  = new RelayCommand(_ => OpenWorkPicker(TomorrowWorkItems));

        AddEquipmentCommand      = new RelayCommand(_ => AddEquipmentRow());
        AddPersonnelCommand      = new RelayCommand(_ => AddPersonnelRow(Personnel));
        AddSujangjaeCommand      = new RelayCommand(_ => AddPersonnelRow(SujangjaePersonnel));

        OpenAiAssistantCommand   = new RelayCommand(_ => OpenAiAssistant());

        InitDefaultEquipment();
        InitDefaultPersonnel();

        // 항목 변경 시 합계 갱신
        Equipment.CollectionChanged          += (_, _) => RefreshEquipmentSums();
        Personnel.CollectionChanged          += (_, _) => RefreshPersonnelSums();
        SujangjaePersonnel.CollectionChanged += (_, _) => RefreshPersonnelSums();

        LoadYesterdayWorkItems();
        LoadYesterdayData();
        LoadDocumentDueList();
        LoadPhotoSlots();
        LoadDocInputCards();
    }

    // ── 기본 장비·인원 초기화 ────────────────────────────────────────

    private void InitDefaultEquipment()
    {
        AddEquipment("B/H 0.2", "5.8t");
        AddEquipment("B/H 380", "39.6t");
        AddEquipment("B/H 480", "64.6t");
        AddEquipment("B/H 1.0", "32.0t");
    }

    private void InitDefaultPersonnel()
    {
        AddPersonnel(Personnel, "관리자");
        AddPersonnel(Personnel, "벌목");
        AddPersonnel(Personnel, "비계");
        AddPersonnel(Personnel, "지하수");
        AddPersonnel(Personnel, "석면");
        AddPersonnel(Personnel, "장비");
        AddPersonnel(SujangjaePersonnel, "수장재");
    }

    // ── 장비 행 추가/연결 ────────────────────────────────────────────

    private EquipmentEntryViewModel AddEquipment(string name = "", string spec = "")
    {
        var vm = new EquipmentEntryViewModel(e =>
        {
            Equipment.Remove(e);
            OnPropertyChanged(nameof(TotalTodayEquipment));
        })
        {
            Name = name,
            Spec = spec,
        };
        vm.PropertyChanged += (_, _) => OnPropertyChanged(nameof(TotalTodayEquipment));
        Equipment.Add(vm);
        return vm;
    }

    private void AddEquipmentRow() => AddEquipment();

    // ── 인원 행 추가/연결 ────────────────────────────────────────────

    private PersonnelEntryViewModel AddPersonnel(
        ObservableCollection<PersonnelEntryViewModel> collection,
        string role = "")
    {
        var vm = new PersonnelEntryViewModel(p =>
        {
            collection.Remove(p);
            RefreshPersonnelSums();
        })
        {
            Role = role,
        };
        vm.PropertyChanged += (_, _) => RefreshPersonnelSums();
        collection.Add(vm);
        return vm;
    }

    private void AddPersonnelRow(ObservableCollection<PersonnelEntryViewModel> collection)
        => AddPersonnel(collection);

    // ── 합계 갱신 ────────────────────────────────────────────────────

    private void RefreshEquipmentSums()
        => OnPropertyChanged(nameof(TotalTodayEquipment));

    private void RefreshPersonnelSums()
    {
        OnPropertyChanged(nameof(TotalTodayPersonnel));
        OnPropertyChanged(nameof(TotalTodayPersonnelWithSujangjae));
    }

    // ── 작업항목 선택 팝업 ────────────────────────────────────────────

    private static void OpenWorkPicker(ObservableCollection<WorkItemViewModel> target)
    {
        var win = new WorkItemPickerWindow { Owner = Application.Current.MainWindow };
        if (win.ShowDialog() == true)
        {
            foreach (var text in win.SelectedItems)
                target.Add(new WorkItemViewModel(text, item => target.Remove(item)));
        }
    }

    // ── AI 도우미 ─────────────────────────────────────────────────────

    private void OpenAiAssistant()
    {
        var win = new AiAssistantWindow { Owner = Application.Current.MainWindow };
        win.SetContext(BuildContextSummary());
        win.Show();
    }

    private string BuildContextSummary()
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"날짜: {TargetDate:yyyy년 MM월 dd일}  날씨: {Weather}");

        if (TodayMorningItems.Any())
            sb.AppendLine("오전 작업: " + string.Join(", ", TodayMorningItems.Select(i => i.Text)));
        if (TodayAfternoonItems.Any())
            sb.AppendLine("오후 작업: " + string.Join(", ", TodayAfternoonItems.Select(i => i.Text)));

        var eqList = Equipment.Where(e => e.TodayCount > 0)
                               .Select(e => $"{e.Name}({e.TodayCount}대)");
        if (eqList.Any())
            sb.AppendLine("장비: " + string.Join(", ", eqList));

        int totalPn = TotalTodayPersonnel + SujangjaePersonnel.Sum(p => p.TodayCount);
        sb.AppendLine($"인원: 공사 {TotalTodayPersonnel}명  수장재 {SujangjaePersonnel.Sum(p => p.TodayCount)}명  합계 {totalPn}명");

        return sb.ToString();
    }

    // ── 문서 목록 로드 ────────────────────────────────────────────────

    private void LoadDocumentDueList()
    {
        DocumentDueList.Clear();
        var    config  = AppConfigLoader.Load();
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        foreach (var doc in config.Documents)
            DocumentDueList.Add(new DocumentDueItem(doc, TargetDate.Date, config.OutputFolder, baseDir));
    }

    // ── 파일 생성 ────────────────────────────────────────────────────

    private void OnGenerate()
    {
        try
        {
            var    config  = AppConfigLoader.Load();
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            var checkedDocs = DocumentDueList.Where(d => d.IsChecked).ToList();
            if (checkedDocs.Count == 0)
            {
                MessageBox.Show("생성할 문서를 선택하세요.", "알림",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var input = BuildDailyInput();

            var dailyDoc = config.Documents.FirstOrDefault(d => d.Id == "daily_report");
            if (dailyDoc != null)
            {
                DateTime yesterday    = TargetDate.AddDays(-1);
                string   prevFileName = GenericExcelWriter.ApplyDatePattern(
                                            dailyDoc.OutputPattern, yesterday);
                string   prevPath     = Path.Combine(baseDir, config.OutputFolder, prevFileName);
                GenericExcelWriter.LoadPreviousDay(prevPath, input, dailyDoc);
            }

            foreach (var card in DocInputCards)
                card.FillInput(input);

            var results = new List<string>();
            var notes   = new List<string>();
            foreach (var dueItem in checkedDocs)
            {
                string outputPath = GenericExcelWriter.Generate(
                    input, dueItem.Config, baseDir, config.OutputFolder,
                    PhotoService.ResizeToCover, notes);
                results.Add($"{dueItem.Name}: {Path.GetFileName(outputPath)}");
            }

            InsertPhotos(checkedDocs, config, baseDir, results);

            if (notes.Count > 0)
                results.Add("\n참고:\n" + string.Join("\n", notes.Take(15))
                            + (notes.Count > 15 ? $"\n… 외 {notes.Count - 15}건" : ""));

            string summary = string.Join("\n", results.Prepend("파일 생성 완료!\n"));
            MessageBox.Show(summary, "완료", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"파일 생성 중 오류가 발생했습니다.\n\n{ex.Message}", "오류",
                            MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private DailyInput BuildDailyInput() => new()
    {
        Date                 = TargetDate,
        Weather              = Weather,
        TodayMorningWork     = string.Join("\n", TodayMorningItems.Select(i => i.Text)),
        TodayAfternoonWork   = string.Join("\n", TodayAfternoonItems.Select(i => i.Text)),
        TomorrowWork         = string.Join("\n", TomorrowWorkItems.Select(i => i.Text)),
        Equipment            = Equipment.Select(e => new EquipmentEntry
        {
            Name         = e.Name,
            Spec         = e.Spec,
            PrevDayCount = e.PrevDayCount,
            TodayCount   = e.TodayCount,
        }).ToList(),
        Personnel            = Personnel.Select(p => new PersonnelEntry
        {
            Role                = p.Role,
            PrevCumulativeCount = p.PrevCumulativeCount,
            TodayCount          = p.TodayCount,
        }).ToList(),
        SujangjaePersonnel   = SujangjaePersonnel.Select(p => new PersonnelEntry
        {
            Role                = p.Role,
            PrevCumulativeCount = p.PrevCumulativeCount,
            TodayCount          = p.TodayCount,
        }).ToList(),
    };

    // ── 작업사항 초기화 ───────────────────────────────────────────────

    private void LoadYesterdayWorkItems()
    {
        var config   = AppConfigLoader.Load();
        var dailyDoc = config.Documents.FirstOrDefault(d => d.Id == "daily_report");
        if (dailyDoc == null) { InitDefaultItems(); return; }

        DateTime yesterday = DateTime.Today.AddDays(-1);
        string   fileName  = GenericExcelWriter.ApplyDatePattern(dailyDoc.OutputPattern, yesterday);
        string   filePath  = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, config.OutputFolder, fileName);

        if (!File.Exists(filePath)) { InitDefaultItems(); return; }

        try
        {
            var todayField    = dailyDoc.Fields.FirstOrDefault(f => f.Id == "today_work");
            var tomorrowField = dailyDoc.Fields.FirstOrDefault(f => f.Id == "tomorrow_work");

            using var doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(
                filePath, isEditable: false);
            var wbPart = doc.WorkbookPart!;
            var wsPart = CellHelper.FindWorksheetPart(wbPart, dailyDoc.SheetName);
            if (wsPart == null) { InitDefaultItems(); return; }

            LoadItemsFromCells(TodayMorningItems, wsPart, todayField?.Cells);
            LoadItemsFromCells(TomorrowWorkItems, wsPart, tomorrowField?.Cells);
        }
        catch { InitDefaultItems(); }
    }

    private void LoadItemsFromCells(
        ObservableCollection<WorkItemViewModel> items,
        DocumentFormat.OpenXml.Packaging.WorksheetPart wsPart,
        List<string>? cells)
    {
        items.Clear();
        if (cells == null) { items.Add(CreateWorkItem("- ", items)); return; }

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
        if (!hasAny) items.Add(CreateWorkItem("- ", items));
    }

    private void InitDefaultItems()
    {
        TodayMorningItems.Clear();
        TodayMorningItems.Add(CreateWorkItem("- ", TodayMorningItems));
        TodayAfternoonItems.Clear();
        TomorrowWorkItems.Clear();
        TomorrowWorkItems.Add(CreateWorkItem("- ", TomorrowWorkItems));
    }

    private static WorkItemViewModel CreateWorkItem(
        string text, ObservableCollection<WorkItemViewModel> owner)
        => new(text, item => owner.Remove(item));

    private static void AddWorkItem(ObservableCollection<WorkItemViewModel> items)
        => items.Add(new WorkItemViewModel("- ", item => items.Remove(item)));

    // ── 어제 누계 로드 ────────────────────────────────────────────────

    private void LoadYesterdayData()
    {
        var config   = AppConfigLoader.Load();
        var dailyDoc = config.Documents.FirstOrDefault(d => d.Id == "daily_report");
        if (dailyDoc == null) return;

        DateTime yesterday = DateTime.Today.AddDays(-1);
        string   fileName  = GenericExcelWriter.ApplyDatePattern(dailyDoc.OutputPattern, yesterday);
        string   filePath  = Path.Combine(
                                 AppDomain.CurrentDomain.BaseDirectory,
                                 config.OutputFolder, fileName);
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
        catch { /* 읽기 오류 시 0 유지 */ }
    }

    private static string NextNextColumn(string col)
    {
        uint idx = CellHelper.ColLetterToIndex(col);
        return ColIndexToLetter(idx + 2);
    }

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

    private static string Norm(string s) => s.Trim().ToLowerInvariant();

    // ── 사진 슬롯 ────────────────────────────────────────────────────

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

    // ── 사진 삽입 ────────────────────────────────────────────────────

    private void InsertPhotos(
        List<DocumentDueItem> checkedDocs,
        AppConfig config,
        string baseDir,
        List<string> results)
    {
        var filledSlots = PhotoSlots.Where(s => s.HasPhoto).ToList();
        if (filledSlots.Count == 0) return;

        foreach (var slot in filledSlots)
        {
            var docConfig = config.Documents.FirstOrDefault(d => d.Id == slot.DocumentId);
            if (docConfig == null) continue;
            var dueItem = checkedDocs.FirstOrDefault(d => d.Config.Id == slot.DocumentId);
            if (dueItem == null) continue;

            string outputFileName = GenericExcelWriter.ApplyDatePattern(
                docConfig.OutputPattern, TargetDate);
            string outputPath = Path.Combine(baseDir, config.OutputFolder, outputFileName);
            if (!File.Exists(outputPath)) continue;

            try
            {
                byte[] imageBytes = PhotoService.ResizeToFill(
                    slot.SelectedImagePath!, slot.TargetWidthPx, slot.TargetHeightPx);
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
