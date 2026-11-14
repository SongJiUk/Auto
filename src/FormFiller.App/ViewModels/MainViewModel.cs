using System.Collections.ObjectModel;
using System.ComponentModel;
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
        {
            var card = new DocInputCardViewModel(doc, () => TargetDate);
            foreach (var (_, item) in card.AllItems) Track(item);
            DocInputCards.Add(card);
        }
        OnPropertyChanged(nameof(HasDocInputs));
        RefreshUsage();
    }

    // ── 단계 (왼쪽 목록) ─────────────────────────────────────────────
    // 오른쪽에는 고른 단계의 입력만 보여 준다. 오늘 만들 문서에 필요 없는 단계는 숨긴다.

    public ObservableCollection<StepItem> Steps { get; } =
    [
        new("docs",   "문서 · 기본 정보"),
        new("work",   "작업사항"),
        new("crew",   "장비 · 인원"),
        new("photos", "사진 · 측정값"),
    ];

    private StepItem Step(string key) => Steps.First(s => s.Key == key);

    private StepItem? _currentStep;
    public StepItem? CurrentStep
    {
        get => _currentStep;
        set
        {
            // ListBox 가 목록을 다시 그릴 때 잠깐 null 을 넣는 경우가 있어 무시한다
            if (value == null || !SetField(ref _currentStep, value)) return;
            OnPropertyChanged(nameof(ShowDocsStep));
            OnPropertyChanged(nameof(ShowWorkStep));
            OnPropertyChanged(nameof(ShowCrewStep));
            OnPropertyChanged(nameof(ShowPhotosStep));
            OnPropertyChanged(nameof(HasPrevStep));
            OnPropertyChanged(nameof(HasNextStep));
        }
    }

    public bool ShowDocsStep   => CurrentStep?.Key == "docs";
    public bool ShowWorkStep   => CurrentStep?.Key == "work";
    public bool ShowCrewStep   => CurrentStep?.Key == "crew";
    public bool ShowPhotosStep => CurrentStep?.Key == "photos";

    private int CurrentStepIndex => CurrentStep == null ? -1 : VisibleSteps.IndexOf(CurrentStep);
    private List<StepItem> VisibleSteps => Steps.Where(s => s.IsVisible).ToList();
    public bool HasPrevStep => CurrentStepIndex > 0;
    public bool HasNextStep => CurrentStepIndex >= 0 && CurrentStepIndex < VisibleSteps.Count - 1;

    public ICommand PrevStepCommand { get; }
    public ICommand NextStepCommand { get; }

    private void MoveStep(int delta)
    {
        var list = VisibleSteps;
        int i = Math.Clamp(CurrentStepIndex + delta, 0, list.Count - 1);
        CurrentStep = list[i];
    }

    // ── 오늘 만들 문서에 필요한 입력 ─────────────────────────────────
    // 체크한 문서의 필드 종류를 보고, 쓰지 않는 입력 칸은 숨긴다.

    private bool _usesWork, _usesEquipment, _usesPersonnel, _usesSujangjae;
    private bool _hasVisibleDocInputs, _hasVisiblePhotoSlots, _noDocChecked;

    public bool UsesWork             => _usesWork;
    public bool UsesEquipment        => _usesEquipment;
    public bool UsesPersonnel        => _usesPersonnel;
    public bool UsesSujangjae        => _usesSujangjae;
    public bool HasVisibleDocInputs  => _hasVisibleDocInputs;
    public bool HasVisiblePhotoSlots => _hasVisiblePhotoSlots;
    public bool NoDocChecked         => _noDocChecked;

    private static bool Has(DocumentConfig d, params string[] types) => d.Fields.Any(f => types.Contains(f.Type));

    private static bool UsesSujangjaeFor(DocumentConfig d) =>
        Has(d, FieldTypes.TbmPersonnel) || (d.IncludesSujangjae && Has(d, FieldTypes.PersonnelTable));

    /// <summary>문서 체크가 바뀌면: 보일 입력 칸과 단계를 다시 정한다.</summary>
    private void RefreshUsage()
    {
        var docs = DocumentDueList.Where(d => d.IsChecked).Select(d => d.Config).ToList();
        var ids  = docs.Select(d => d.Id).ToHashSet();

        SetField(ref _usesWork,      docs.Any(d => Has(d, FieldTypes.Multiline)), nameof(UsesWork));
        SetField(ref _usesEquipment, docs.Any(d => Has(d, FieldTypes.EquipmentTable, FieldTypes.ActiveEquipmentNames)), nameof(UsesEquipment));
        SetField(ref _usesPersonnel, docs.Any(d => Has(d, FieldTypes.PersonnelTable, FieldTypes.TbmPersonnel)), nameof(UsesPersonnel));
        SetField(ref _usesSujangjae, docs.Any(UsesSujangjaeFor), nameof(UsesSujangjae));

        foreach (var card in DocInputCards) card.IsVisible = ids.Contains(card.Config.Id);
        foreach (var slot in PhotoSlots)    slot.IsVisible = ids.Contains(slot.DocumentId);
        SetField(ref _hasVisibleDocInputs,  DocInputCards.Any(c => c.IsVisible), nameof(HasVisibleDocInputs));
        SetField(ref _hasVisiblePhotoSlots, PhotoSlots.Any(s => s.IsVisible),    nameof(HasVisiblePhotoSlots));
        SetField(ref _noDocChecked,         docs.Count == 0,                     nameof(NoDocChecked));

        Step("work").IsVisible   = _usesWork;
        Step("crew").IsVisible   = _usesEquipment || _usesPersonnel || _usesSujangjae;
        Step("photos").IsVisible = _hasVisibleDocInputs || _hasVisiblePhotoSlots;

        int n = 1;
        foreach (var s in Steps.Where(s => s.IsVisible)) s.Number = n++;

        if (CurrentStep == null || !CurrentStep.IsVisible) CurrentStep = Steps[0];
        OnPropertyChanged(nameof(HasPrevStep));
        OnPropertyChanged(nameof(HasNextStep));
        RefreshSteps();
    }

    /// <summary>단계마다 ✔(다 했음)와 요약 글자를 다시 계산한다. 입력이 바뀔 때마다 부른다.</summary>
    private void RefreshSteps()
    {
        int docCount = DocumentDueList.Count(d => d.IsChecked);
        var docs = Step("docs");
        docs.IsDone  = docCount > 0;
        docs.Summary = docCount > 0 ? $"{TargetDate:MM/dd (ddd)} · 문서 {docCount}개" : "만들 문서를 고르세요";

        int am = CountFilled(TodayMorningItems), pm = CountFilled(TodayAfternoonItems), tm = CountFilled(TomorrowWorkItems);
        var work = Step("work");
        work.IsDone  = am + pm > 0;
        work.Summary = $"오전 {am} · 오후 {pm} · 명일 {tm}";

        var crew = Step("crew");
        crew.IsDone  = (!_usesEquipment || TotalTodayEquipment > 0)
                    && (!_usesPersonnel || TotalTodayPersonnelWithSujangjae > 0);
        crew.Summary = $"장비 {TotalTodayEquipment}대 · 인원 {TotalTodayPersonnelWithSujangjae}명";

        var items  = DocInputCards.Where(c => c.IsVisible).SelectMany(c => c.AllItems).Select(x => x.Item).ToList();
        var slots  = PhotoSlots.Where(s => s.IsVisible).ToList();
        int total  = items.Count(i => i.IsPhoto) + slots.Count;
        int filled = items.Count(i => i.IsPhoto && i.HasPhoto) + slots.Count(s => s.HasPhoto);
        int noVal  = items.Count(i => i.HasValue && i.Value.Trim().Length == 0);
        var photos = Step("photos");
        photos.IsDone  = filled == total && noVal == 0;
        photos.Summary = (total > 0 ? $"사진 {filled}/{total}" : "")
                       + (noVal > 0 ? (total > 0 ? " · " : "") + $"값 {noVal}칸 남음" : "");
    }

    private static int CountFilled(IEnumerable<WorkItemViewModel> items) => items.Count(i => !IsBlankWork(i.Text));

    /// <summary>"- " 처럼 내용 없는 작업 줄인지</summary>
    private static bool IsBlankWork(string text) => text.Trim().TrimStart('-').Trim().Length == 0;

    /// <summary>항목 값이 바뀌면 단계 ✔ 를 다시 계산하도록 연결한다.</summary>
    private void Track(INotifyPropertyChanged item) => item.PropertyChanged += (_, _) => RefreshSteps();

    private void TrackCollection<T>(ObservableCollection<T> list) where T : INotifyPropertyChanged
    {
        list.CollectionChanged += (_, e) =>
        {
            if (e.NewItems != null)
                foreach (T item in e.NewItems) Track(item);
            RefreshSteps();
        };
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

        PrevStepCommand          = new RelayCommand(_ => MoveStep(-1), _ => HasPrevStep);
        NextStepCommand          = new RelayCommand(_ => MoveStep(+1), _ => HasNextStep);

        // 입력이 바뀌면 왼쪽 단계의 ✔ 와 요약을 갱신 (항목을 넣기 전에 연결해 둔다)
        TrackCollection(TodayMorningItems);
        TrackCollection(TodayAfternoonItems);
        TrackCollection(TomorrowWorkItems);
        TrackCollection(Equipment);
        TrackCollection(Personnel);
        TrackCollection(SujangjaePersonnel);

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
        {
            var item = new DocumentDueItem(doc, TargetDate.Date, config.OutputFolder, baseDir);
            item.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(DocumentDueItem.IsChecked)) RefreshUsage();
            };
            DocumentDueList.Add(item);
        }
        RefreshUsage();
    }

    // ── 파일 생성 ────────────────────────────────────────────────────
    // [파일 생성] → 확인 창(빠진 칸·문제) → [생성] → 같은 창에 결과와 [열기] 버튼

    private void OnGenerate()
    {
        var    config  = AppConfigLoader.Load();
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        var checkedDocs = DocumentDueList.Where(d => d.IsChecked).ToList();
        if (checkedDocs.Count == 0)
        {
            MessageBox.Show("만들 문서를 하나 이상 체크하세요.", "알림",
                            MessageBoxButton.OK, MessageBoxImage.Information);
            CurrentStep = Step("docs");
            return;
        }

        var checks = checkedDocs.Select(d => BuildCheck(d.Config, config, baseDir)).ToList();

        // 저장 파일 이름이 같은 문서끼리는 나중 것이 앞의 것을 덮어쓴다
        foreach (var group in checks.GroupBy(c => c.OutputPath, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
            foreach (var c in group)
                c.Warn($"다른 문서({string.Join(", ", group.Where(o => o != c).Select(o => o.Name))})와 저장 파일 이름이 같아 서로 덮어씁니다. 설정에서 파일 이름을 바꿔 주세요.");
        var vm = new GenerateViewModel(TargetDate, checks, Path.Combine(baseDir, config.OutputFolder),
                                       targets => RunGenerate(targets, config, baseDir));
        new GenerateWindow(vm) { Owner = Application.Current.MainWindow }.ShowDialog();
    }

    /// <summary>
    /// 문서 하나를 만들기 전에 확인할 것을 모은다.
    /// Error(빨강) = 만들 수 없음 → 그 문서는 건너뜀. Warn(주황) = 만들 수는 있지만 확인 필요.
    /// </summary>
    private DocCheckItem BuildCheck(DocumentConfig doc, AppConfig config, string baseDir)
    {
        var c = new DocCheckItem(doc, GenericExcelWriter.BuildOutputPath(
                                          TargetDate, doc.OutputPattern, baseDir, config.OutputFolder));

        // 서식 파일
        if (string.IsNullOrWhiteSpace(doc.TemplateFile) || !File.Exists(Path.Combine(baseDir, doc.TemplateFile)))
            c.Error("서식 파일이 없습니다. 설정에서 서식 파일을 다시 골라 주세요.");

        // 같은 이름 파일
        if (File.Exists(c.OutputPath))
        {
            if (IsFileLocked(c.OutputPath))
                c.Error("같은 날짜 파일이 Excel에서 열려 있어 저장할 수 없습니다. Excel에서 닫고 다시 누르세요.");
            else
                c.Warn("이 날짜 파일이 이미 있습니다. 새로 만들면 덮어씁니다.");
        }

        // 작업사항
        if (Has(doc, FieldTypes.Multiline))
        {
            if (CountFilled(TodayMorningItems) + CountFilled(TodayAfternoonItems) == 0)
                c.Warn("금일 작업사항이 비어 있습니다.");
            if (CountFilled(TomorrowWorkItems) == 0 && doc.Fields.Any(f => f.Id == "tomorrow_work"))
                c.Warn("명일 작업사항이 비어 있습니다.");
        }

        // 장비
        if (Has(doc, FieldTypes.EquipmentTable, FieldTypes.ActiveEquipmentNames))
        {
            if (TotalTodayEquipment == 0)
                c.Warn("장비 금일 대수가 모두 0입니다.");
            if (Equipment.Any(e => e.TodayCount > 0 && string.IsNullOrWhiteSpace(e.Name)))
                c.Warn("이름 없는 장비 줄에 대수가 들어 있습니다.");
        }

        // 인원
        if (Has(doc, FieldTypes.PersonnelTable, FieldTypes.TbmPersonnel))
        {
            int people = UsesSujangjaeFor(doc) ? TotalTodayPersonnelWithSujangjae : TotalTodayPersonnel;
            if (people == 0) c.Warn("금일 인원이 0명입니다.");
        }

        // 전일 값: 공사일보는 생성할 때 어제 파일에서 다시 읽는다. 파일이 없으면 화면 값을 그대로 쓴다.
        if (doc.Id == "daily_report" && Has(doc, FieldTypes.EquipmentTable, FieldTypes.PersonnelTable))
        {
            string prev = GenericExcelWriter.BuildOutputPath(
                TargetDate.AddDays(-1), doc.OutputPattern, baseDir, config.OutputFolder);
            if (!File.Exists(prev))
                c.Warn($"어제 파일({Path.GetFileName(prev)})이 없어 화면의 전일 값을 그대로 씁니다. 전일·누계를 확인하세요.");
        }

        // 사진·측정값 (시트마다 한 줄로 묶어서)
        var card = DocInputCards.FirstOrDefault(x => x.Config.Id == doc.Id);
        if (card != null)
        {
            foreach (var sheet in card.Sheets)
            {
                var noPhoto = sheet.Items.Where(i => i.IsPhoto && !i.HasPhoto).Select(i => i.Label).ToList();
                var noValue = sheet.Items.Where(i => i.HasValue && i.Value.Trim().Length == 0).Select(i => i.Label).ToList();
                string where = card.Sheets.Count > 1 ? $"[{sheet.Header}] " : "";
                if (noPhoto.Count > 0)
                    c.Warn($"{where}사진 없음 {noPhoto.Count}칸 ({ShortList(noPhoto)}) → 서식 사진 그대로");
                if (noValue.Count > 0)
                    c.Warn($"{where}값 없음 {noValue.Count}칸 ({ShortList(noValue)}) → 서식 값 그대로");
            }
        }

        int emptySlots = PhotoSlots.Count(s => s.DocumentId == doc.Id && !s.HasPhoto);
        if (emptySlots > 0)
            c.Warn($"사진대지 빈 칸 {emptySlots}개 → 서식 사진 그대로");

        return c;
    }

    private static string ShortList(List<string> labels) =>
        string.Join(", ", labels.Take(3)) + (labels.Count > 3 ? " …" : "");

    /// <summary>다른 프로그램(Excel)이 파일을 열고 있어 덮어쓸 수 없는지</summary>
    private static bool IsFileLocked(string path)
    {
        try
        {
            using var fs = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return false;
        }
        catch (IOException)                  { return true; }
        catch (UnauthorizedAccessException)  { return true; }
    }

    /// <summary>확인 창에서 [생성] 을 누르면 호출된다. 문서마다 따로 만들어, 하나가 실패해도 나머지는 만든다.</summary>
    private List<string> RunGenerate(IReadOnlyList<DocCheckItem> targets, AppConfig config, string baseDir)
    {
        var notes = new List<string>();
        var input = BuildDailyInput();

        var dailyDoc = config.Documents.FirstOrDefault(d => d.Id == "daily_report");
        if (dailyDoc != null)
        {
            string prevPath = GenericExcelWriter.BuildOutputPath(
                TargetDate.AddDays(-1), dailyDoc.OutputPattern, baseDir, config.OutputFolder);
            GenericExcelWriter.LoadPreviousDay(prevPath, input, dailyDoc);
        }

        foreach (var card in DocInputCards)
            card.FillInput(input);

        foreach (var t in targets)
        {
            try
            {
                GenericExcelWriter.Generate(input, t.Config, baseDir, config.OutputFolder,
                                            PhotoService.ResizeToCover, notes);
                t.SetDone();
            }
            catch (Exception ex)
            {
                t.SetFailed(ex.Message);
            }
        }

        var doneIds = targets.Where(t => t.IsOk).Select(t => t.Config.Id).ToHashSet();
        InsertPhotos(doneIds, config, baseDir, notes);
        return notes;
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
                Track(slot);
                PhotoSlots.Add(slot);
            }
        }
        OnPropertyChanged(nameof(HasPhotoSlots));
        RefreshUsage();
    }

    // ── 사진 삽입 ────────────────────────────────────────────────────

    private void InsertPhotos(
        HashSet<string> generatedDocIds,
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
            if (!generatedDocIds.Contains(slot.DocumentId)) continue;

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
                results.Add($"사진대지 사진 넣음: {slot.Label} → {slot.SheetName}");
            }
            catch (Exception ex)
            {
                results.Add($"사진대지 사진 넣기 실패({slot.Label}): {ex.Message}");
            }
        }
    }
}
