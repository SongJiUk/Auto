using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using FormFiller.Core.Config;
using FormFiller.Core.Models;

namespace FormFiller.App.ViewModels;

// ════════════════════════════════════════════════════════════════════
// 메인 화면 "항목 입력" 단계: 블록 하나 = 입력 카드 하나.
// 같은 이름(Key)의 블록은 여러 문서에 있어도 카드 하나 (한 번 입력하면 모든 문서에 들어간다).
// 카드 모양은 MainWindow.xaml 의 DataTemplate 이 종류(클래스)마다 정한다.
//   (DataTemplate DataType="{x:Type vm:ValueInputViewModel}" → 그 클래스면 그 모양으로 그린다)
// ════════════════════════════════════════════════════════════════════

/// <summary>입력 카드 공통</summary>
public abstract class BlockInputViewModel : ViewModelBase
{
    public string Key { get; }
    public FieldConfig Field { get; }

    /// <summary>이 값을 쓰는 문서 이름들 (카드 아래 작은 글씨)</summary>
    public List<string> DocNames { get; } = [];
    public HashSet<string> DocIds { get; } = [];
    public string UsedIn => "쓰는 문서: " + string.Join(", ", DocNames);

    public abstract string Icon { get; }
    public string Title => Key;

    private bool _isVisible = true;
    /// <summary>오늘 만들 문서(체크)에 이 블록이 있을 때만 보인다</summary>
    public bool IsVisible { get => _isVisible; set => SetField(ref _isVisible, value); }

    /// <summary>사용자가 값을 바꾸면 (단계 ✔·닫을 때 확인·계산 미리보기 갱신용)</summary>
    public event Action<BlockInputViewModel>? Edited;
    protected void RaiseEdited() => Edited?.Invoke(this);

    /// <summary>사용자가 직접 고쳤는지 (날짜를 바꿔 처음 값을 다시 읽을 때 덮어쓰지 않으려고)</summary>
    public bool IsEdited { get; protected set; }

    protected BlockInputViewModel(string key, FieldConfig field) { Key = key; Field = field; }

    /// <summary>어제 파일(없으면 서식)에서 만든 처음 값을 넣는다. keepEdits: 이미 입력한 값은 살린다.</summary>
    public abstract void Load(BlockValue initial, bool keepEdits);

    /// <summary>지금 화면 값 → 엔진 값</summary>
    public abstract BlockValue ToValue();

    /// <summary>다 채웠는지 (단계 ✔)</summary>
    public virtual bool IsFilled => true;

    public void AddDoc(DocumentConfig doc)
    {
        if (!DocIds.Add(doc.Id)) return;
        DocNames.Add(doc.Name);
        OnPropertyChanged(nameof(UsedIn));
    }

    public static BlockInputViewModel? Create(string key, FieldConfig f) => f.Type switch
    {
        BlockTypes.Value when (f.ValueSource ?? BlockTypes.SourceInput) == BlockTypes.SourceInput => new ValueInputViewModel(key, f),
        BlockTypes.Lines   => new LinesInputViewModel(key, f),
        BlockTypes.Counter => new CounterInputViewModel(key, f),
        BlockTypes.Table   => new TableInputViewModel(key, f),
        BlockTypes.Roster  => new RosterInputViewModel(key, f),
        BlockTypes.Calc    => new CalcInputViewModel(key, f),
        BlockTypes.Photos  => new PhotosInputViewModel(key, f),
        _                  => null,   // 날짜·고정값·어제 값 그대로 → 입력할 것 없음
    };
}

// ── 값 ──────────────────────────────────────────────────────────────

public class ValueInputViewModel : BlockInputViewModel
{
    public override string Icon => "🔤";
    public string Suffix => Field.Suffix ?? "";
    public bool HasSuffix => Suffix.Length > 0;

    private string _text = "";
    public string Text
    {
        get => _text;
        set { if (SetField(ref _text, value)) { IsEdited = true; RaiseEdited(); } }
    }

    public ValueInputViewModel(string key, FieldConfig f) : base(key, f) { }

    public override void Load(BlockValue initial, bool keepEdits)
    {
        if (keepEdits && IsEdited) return;
        _text = initial.Text;
        OnPropertyChanged(nameof(Text));
    }

    public override BlockValue ToValue() => new() { Text = Text };
    public override bool IsFilled => Text.Trim().Length > 0;
}

// ── 여러 줄 ─────────────────────────────────────────────────────────

public class LinesInputViewModel : BlockInputViewModel
{
    public override string Icon => "📝";
    public ObservableCollection<WorkItemViewModel> Lines { get; } = [];
    public ICommand AddLineCommand { get; }
    public ICommand PickCommand { get; }
    public string Prefix => Field.Prefix ?? "";

    public LinesInputViewModel(string key, FieldConfig f) : base(key, f)
    {
        AddLineCommand = new RelayCommand(_ => { Add(Prefix); IsEdited = true; RaiseEdited(); });
        PickCommand    = new RelayCommand(_ => PickRequested?.Invoke(this));
        Lines.CollectionChanged += (_, _) => RaiseEdited();
    }

    /// <summary>[☰ 항목 선택] — 자주 쓰는 작업 목록 창은 MainViewModel 이 연다</summary>
    public event Action<LinesInputViewModel>? PickRequested;

    public void Add(string text)
    {
        var item = new WorkItemViewModel(text, i => { Lines.Remove(i); IsEdited = true; });
        item.PropertyChanged += (_, e) => { if (e.PropertyName == nameof(WorkItemViewModel.Text)) { IsEdited = true; RaiseEdited(); } };
        Lines.Add(item);
    }

    public override void Load(BlockValue initial, bool keepEdits)
    {
        if (keepEdits && IsEdited) return;
        Lines.Clear();
        foreach (var l in initial.Lines) Add(l);
        if (Lines.Count == 0) Add(Prefix);
        IsEdited = false;
    }

    public override BlockValue ToValue() => new() { Lines = Lines.Select(l => l.Text).ToList() };
    public override bool IsFilled => Lines.Any(l => l.Text.Trim().TrimStart('-').Trim().Length > 0);
}

// ── 숫자 누계 ───────────────────────────────────────────────────────

public class CounterInputViewModel : BlockInputViewModel
{
    public override string Icon => "➕";

    private decimal _prev, _today;
    /// <summary>어제까지 누계 (어제 파일에서 읽음, 고칠 수 있음)</summary>
    public decimal Prev
    {
        get => _prev;
        set { if (SetField(ref _prev, value)) { OnPropertyChanged(nameof(Total)); IsEdited = true; RaiseEdited(); } }
    }
    public decimal Today
    {
        get => _today;
        set { if (SetField(ref _today, value)) { OnPropertyChanged(nameof(Total)); IsEdited = true; RaiseEdited(); } }
    }
    public decimal Total => Prev + Today;

    public CounterInputViewModel(string key, FieldConfig f) : base(key, f) { }

    public override void Load(BlockValue initial, bool keepEdits)
    {
        _prev = initial.Prev;                       // 전일은 날짜가 바뀌면 항상 다시 읽는다
        if (!(keepEdits && IsEdited)) _today = 0;
        OnPropertyChanged(nameof(Prev)); OnPropertyChanged(nameof(Today)); OnPropertyChanged(nameof(Total));
    }

    public override BlockValue ToValue() => new() { Prev = Prev, Today = Today };
}

// ── 표 ──────────────────────────────────────────────────────────────

public class TableRowInputViewModel : ViewModelBase
{
    private string _name = "", _spec = "";
    private decimal _prev, _today;
    public string Name { get => _name; set => SetField(ref _name, value); }
    public string Spec { get => _spec; set => SetField(ref _spec, value); }
    public decimal Prev  { get => _prev;  set { if (SetField(ref _prev, value))  OnPropertyChanged(nameof(Total)); } }
    public decimal Today { get => _today; set { if (SetField(ref _today, value)) OnPropertyChanged(nameof(Total)); } }
    public decimal Total => Prev + Today;
    public decimal LastToday { get; set; }
    public ICommand RemoveCommand { get; }
    public TableRowInputViewModel(Action<TableRowInputViewModel> onRemove) => RemoveCommand = new RelayCommand(_ => onRemove(this));
}

public class TableInputViewModel : BlockInputViewModel
{
    public override string Icon => "📊";
    public ObservableCollection<TableRowInputViewModel> Rows { get; } = [];
    public ICommand AddRowCommand { get; }
    public ICommand CopyLastTodayCommand { get; }

    public bool HasSpec => BlockTypes.ColOf(Field, BlockTypes.RoleSpec) != null;
    /// <summary>금일을 명단에서 센다 (금일 칸은 읽기 전용)</summary>
    public bool FromRoster => !string.IsNullOrWhiteSpace(Field.RosterKey);
    public bool TodayEditable => !FromRoster;
    public string RosterNote => FromRoster ? $"금일은 '{Field.RosterKey}' 명단에서 직종별로 자동으로 셉니다." : "";

    public decimal TodaySum => Rows.Sum(r => r.Today);

    public TableInputViewModel(string key, FieldConfig f) : base(key, f)
    {
        AddRowCommand        = new RelayCommand(_ => { AddRow(new TableRowValue()); IsEdited = true; });
        CopyLastTodayCommand = new RelayCommand(_ => { foreach (var r in Rows) r.Today = r.LastToday; }, _ => TodayEditable && Rows.Any(r => r.LastToday > 0));
        Rows.CollectionChanged += (_, _) => { OnPropertyChanged(nameof(TodaySum)); RaiseEdited(); };
    }

    private TableRowInputViewModel AddRow(TableRowValue v)
    {
        var row = new TableRowInputViewModel(r => { Rows.Remove(r); IsEdited = true; })
            { Name = v.Name, Spec = v.Spec, Prev = v.Prev, Today = v.Today, LastToday = v.LastToday };
        row.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName is nameof(TableRowInputViewModel.Total) or nameof(TableRowInputViewModel.LastToday)) return;
            if (e.PropertyName == nameof(TableRowInputViewModel.Today)) OnPropertyChanged(nameof(TodaySum));
            IsEdited = true;
            RaiseEdited();
        };
        Rows.Add(row);
        return row;
    }

    /// <summary>
    /// 줄 목록은 어제 파일 기준으로 다시 만들고, 이미 입력한 금일(과 새로 추가한 줄)은 이름으로 찾아 살린다.
    /// </summary>
    public override void Load(BlockValue initial, bool keepEdits)
    {
        var old = keepEdits ? Rows.Select(r => (r.Name, r.Spec, r.Today)).ToList() : [];
        Rows.Clear();
        foreach (var v in initial.Rows)
        {
            var row = AddRow(v);
            int i = old.FindIndex(o => o.Name.Trim() == v.Name.Trim());
            if (i >= 0) { row.Today = old[i].Today; old.RemoveAt(i); }
        }
        foreach (var o in old.Where(o => o.Name.Trim().Length > 0))   // 화면에서 추가했던 줄
            AddRow(new TableRowValue { Name = o.Name, Spec = o.Spec, Today = o.Today });
        IsEdited = false;
    }

    /// <summary>명단에서 센 금일을 넣는다 (FromRoster 일 때 MainViewModel 이 부른다)</summary>
    public void SetCounts(IEnumerable<TableRowValue> counted)
    {
        foreach (var c in counted)
            if (Rows.FirstOrDefault(r => r.Name == c.Name) is { } row) row.Today = c.Today;
    }

    public override BlockValue ToValue() => new()
    {
        Rows = Rows.Where(r => r.Name.Trim().Length > 0)
                   .Select(r => new TableRowValue { Name = r.Name.Trim(), Spec = r.Spec.Trim(), Prev = r.Prev, Today = r.Today, LastToday = r.LastToday })
                   .ToList(),
    };

    public override bool IsFilled => FromRoster || TodaySum > 0;
}

// ── 명단 ────────────────────────────────────────────────────────────

public class PersonRowViewModel : ViewModelBase
{
    private bool _isChecked;
    private string _name = "", _job = "", _task = "";
    /// <summary>오늘 나왔는지</summary>
    public bool IsChecked { get => _isChecked; set => SetField(ref _isChecked, value); }
    public string Name { get => _name; set => SetField(ref _name, value); }
    public string Job  { get => _job;  set => SetField(ref _job, value); }
    public string Task { get => _task; set => SetField(ref _task, value); }
    public ICommand RemoveCommand { get; }
    public PersonRowViewModel(Action<PersonRowViewModel> onRemove) => RemoveCommand = new RelayCommand(_ => onRemove(this));
}

public class RosterInputViewModel : BlockInputViewModel
{
    public override string Icon => "👥";

    /// <summary>사람 목록 (설정에 저장된 사람 + 어제 명단). 체크한 사람만, 위에서부터 순서대로 쓴다.</summary>
    public ObservableCollection<PersonRowViewModel> People { get; } = [];
    public ICommand AddPersonCommand { get; }
    public ICommand CheckAllCommand  { get; }
    public ICommand UncheckAllCommand { get; }

    public int CheckedCount => People.Count(p => p.IsChecked);
    /// <summary>직종별 인원 (예: "장비 10 · 소장 1 · 합계 13")</summary>
    public string CountText => CheckedCount == 0 ? "체크한 사람이 없습니다" :
        string.Join(" · ", People.Where(p => p.IsChecked).GroupBy(p => p.Job.Trim().Length == 0 ? "(직종 없음)" : p.Job.Trim())
                                 .Select(g => $"{g.Key} {g.Count()}")) + $" · 합계 {CheckedCount}명";

    public RosterInputViewModel(string key, FieldConfig f) : base(key, f)
    {
        AddPersonCommand  = new RelayCommand(_ => { var p = AddPerson(new PersonConfig(), true); IsEdited = true; RaiseEdited(); });
        CheckAllCommand   = new RelayCommand(_ => { foreach (var p in People) p.IsChecked = true; });
        UncheckAllCommand = new RelayCommand(_ => { foreach (var p in People) p.IsChecked = false; });
        People.CollectionChanged += (_, _) => Notify();
    }

    private void Notify()
    {
        OnPropertyChanged(nameof(CheckedCount));
        OnPropertyChanged(nameof(CountText));
        RaiseEdited();
    }

    private PersonRowViewModel AddPerson(PersonConfig p, bool isChecked)
    {
        var row = new PersonRowViewModel(r => { People.Remove(r); IsEdited = true; })
            { Name = p.Name, Job = p.Role, Task = p.Task, IsChecked = isChecked };
        row.PropertyChanged += (_, _) => { IsEdited = true; Notify(); };
        People.Add(row);
        return row;
    }

    /// <summary>
    /// 사람 목록 = 어제 명단(체크된 채, 어제 순서) + 설정에 저장된 나머지 사람(체크 안 됨).
    /// keepEdits 면 지금 체크·작업내용을 살린다.
    /// </summary>
    public void Load(BlockValue initial, IReadOnlyList<PersonConfig> saved, bool keepEdits)
    {
        var old = keepEdits && IsEdited ? People.Select(p => (p.Name, p.Job, p.Task, p.IsChecked)).ToList() : null;
        People.Clear();
        if (old != null)
        {
            foreach (var o in old) AddPerson(new PersonConfig { Name = o.Name, Role = o.Job, Task = o.Task }, o.IsChecked);
            return;
        }
        foreach (var e in initial.People)
            AddPerson(new PersonConfig { Name = e.Person, Role = e.Job, Task = e.Task }, true);
        foreach (var p in saved.Where(s => People.All(x => x.Name.Trim() != s.Name.Trim())))
            AddPerson(p, false);
        IsEdited = false;
        Notify();
    }

    public override void Load(BlockValue initial, bool keepEdits) => Load(initial, [], keepEdits);

    public override BlockValue ToValue() => new()
    {
        People = People.Where(p => p.IsChecked && p.Name.Trim().Length > 0)
                       .Select(p => new RosterEntry { Person = p.Name.Trim(), Job = p.Job.Trim(), Task = p.Task.Trim() }).ToList(),
    };

    /// <summary>설정에 저장할 사람 목록 (체크 여부 상관없이 모두)</summary>
    public IEnumerable<PersonConfig> AllPeople() =>
        People.Where(p => p.Name.Trim().Length > 0)
              .Select(p => new PersonConfig { Name = p.Name.Trim(), Role = p.Job.Trim(), Task = p.Task.Trim() });

    public override bool IsFilled => CheckedCount > 0;
}

// ── 사진 구역 ───────────────────────────────────────────────────────

/// <summary>사진 구역 카드의 사진 칸 하나 (사진 + 설명)</summary>
public class PhotoPickViewModel : ViewModelBase
{
    private readonly Action _onEdited;
    public string Label { get; }
    /// <summary>이 칸에 설명 칸이 있는지 (없으면 설명 입력 상자를 숨긴다)</summary>
    public bool HasCaption { get; }

    private string _path = "";
    public string Path
    {
        get => _path;
        set
        {
            if (!SetField(ref _path, value)) return;
            OnPropertyChanged(nameof(HasPhoto));
            OnPropertyChanged(nameof(FileName));
            LoadThumbnail(value);
            _onEdited();
        }
    }
    public bool HasPhoto => _path.Length > 0;
    public string FileName => HasPhoto ? System.IO.Path.GetFileName(_path) : "";

    private string _caption = "";
    public string Caption { get => _caption; set { if (SetField(ref _caption, value)) _onEdited(); } }

    private System.Windows.Media.Imaging.BitmapImage? _thumbnail;
    public System.Windows.Media.Imaging.BitmapImage? Thumbnail { get => _thumbnail; private set => SetField(ref _thumbnail, value); }

    public ICommand BrowseCommand { get; }
    public ICommand ClearCommand  { get; }

    public PhotoPickViewModel(string label, bool hasCaption, Action onEdited)
    {
        Label = label; HasCaption = hasCaption; _onEdited = onEdited;
        BrowseCommand = new RelayCommand(_ =>
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
                { Title = $"{Label} 선택", Filter = "이미지 파일 (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp" };
            if (dlg.ShowDialog() == true) Path = dlg.FileName;
        });
        ClearCommand = new RelayCommand(_ => Path = "", _ => HasPhoto);
    }

    /// <summary>끌어다 놓은 파일 (사진만 받는다)</summary>
    public void SetDropped(string file)
    {
        string ext = System.IO.Path.GetExtension(file).ToLowerInvariant();
        if (ext is ".jpg" or ".jpeg" or ".png" or ".bmp") Path = file;
    }

    /// <summary>처음 값을 넣을 때 (사용자가 고친 것이 아니므로 Edited 를 올리지 않는다)</summary>
    public void Init(string path, string caption)
    {
        _caption = caption; OnPropertyChanged(nameof(Caption));
        if (_path == path) return;
        _path = path;
        OnPropertyChanged(nameof(Path)); OnPropertyChanged(nameof(HasPhoto)); OnPropertyChanged(nameof(FileName));
        LoadThumbnail(path);
    }

    /// <summary>미리보기는 파일을 읽어야 해서(드라이브면 느림) 뒤에서 읽는다</summary>
    private async void LoadThumbnail(string path)
    {
        Thumbnail = null;
        if (path.Length == 0) return;
        try
        {
            var thumb = await BackgroundWork.RunSta(() => PhotoService.LoadThumbnail(path));
            if (_path == path) Thumbnail = thumb;
        }
        catch (Exception ex) { AppLog.Error($"사진 미리보기 실패: {path}", ex); }
    }
}

/// <summary>
/// 사진 구역: 설정한 사진 칸 수만큼 칸을 보여 준다. 그날 넣은 사진만 칸 1부터 차례로 들어가고,
/// 1장이면 설정의 "1장일 때 자리"에 들어간다 (엔진 PhotoAreaWriter).
/// </summary>
public class PhotosInputViewModel : BlockInputViewModel
{
    public override string Icon => "📷";
    public ObservableCollection<PhotoPickViewModel> Slots { get; } = [];
    public ICommand ImportCommand { get; }

    /// <summary>[📂 폴더에서 오늘 사진 불러오기] — 폴더는 문서 설정에 있어서 MainViewModel 이 처리한다</summary>
    public event Action<PhotosInputViewModel>? ImportRequested;

    public string Note => Field.SinglePhotoArea != null
        ? $"사진 칸 {Slots.Count}개 · 1장만 넣으면 '1장일 때 자리'({Field.SinglePhotoArea.Range})에 크게 들어갑니다."
        : $"사진 칸 {Slots.Count}개 · 넣은 사진만 칸 1부터 차례로 들어갑니다. 비우면 어제 사진 그대로.";

    public PhotosInputViewModel(string key, FieldConfig f) : base(key, f)
    {
        var areas = f.PhotoAreas ?? [];
        for (int i = 0; i < Math.Max(1, areas.Count); i++)
        {
            bool hasCaption = (i < areas.Count && !string.IsNullOrEmpty(areas[i].CaptionCell))
                              || (i == 0 && !string.IsNullOrEmpty(f.SinglePhotoArea?.CaptionCell));
            Slots.Add(new PhotoPickViewModel($"사진 {i + 1}", hasCaption, () => { IsEdited = true; RaiseEdited(); }));
        }
        ImportCommand = new RelayCommand(_ => ImportRequested?.Invoke(this));
    }

    /// <summary>사진은 매일 새로 넣고, 설명은 어제 설명을 처음 값으로 보여 준다.</summary>
    public override void Load(BlockValue initial, bool keepEdits)
    {
        if (keepEdits && IsEdited) return;
        for (int i = 0; i < Slots.Count; i++)
            Slots[i].Init("", i < initial.Photos.Count ? initial.Photos[i].Caption : "");
        IsEdited = false;
    }

    /// <summary>폴더에서 찾은 사진을 칸 1부터 넣는다 (넣은 장수를 돌려준다)</summary>
    public int SetPhotos(IReadOnlyList<string> paths)
    {
        int n = Math.Min(paths.Count, Slots.Count);
        for (int i = 0; i < Slots.Count; i++) Slots[i].Path = i < n ? paths[i] : "";
        return n;
    }

    public override BlockValue ToValue() => new()
    {
        Photos = Slots.Select(s => new PhotoValue { Path = s.Path, Caption = s.Caption }).ToList(),
    };

    public override bool IsFilled => Slots.Any(s => s.HasPhoto);
}

// ── 계산 (입력 없음, 미리보기만) ────────────────────────────────────

public class CalcInputViewModel : BlockInputViewModel
{
    public override string Icon => "🧮";
    private string _result = "";
    /// <summary>지금 입력으로 계산한 값 (예: "15")</summary>
    public string Result { get => _result; set => SetField(ref _result, value); }
    public string Formula => (Field.CalcKind ?? BlockTypes.CalcSum) switch
    {
        BlockTypes.CalcNames => $"이름 나열: {string.Join(", ", Field.Sources ?? [])}",
        BlockTypes.CalcCount => $"인원 수: {string.Join(", ", Field.Sources ?? [])}",
        _                    => $"합계: {string.Join(" + ", Field.Sources ?? [])}",
    };
    public CalcInputViewModel(string key, FieldConfig f) : base(key, f) { }
    public override void Load(BlockValue initial, bool keepEdits) { }
    public override BlockValue ToValue() => new();
}
