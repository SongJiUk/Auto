using System.Collections.ObjectModel;
using System.Windows.Input;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;

namespace FormFiller.App.ViewModels.Settings;

// ════════════════════════════════════════════════════════════════════
// 설정 화면 ↔ Excel 연결
// ════════════════════════════════════════════════════════════════════

/// <summary>
/// 설정 화면 ViewModel 이 Excel(왼쪽 미리보기)에 물어보거나 시키는 일.
/// 실제 Excel COM 호출은 View(DocumentSetupWindow)가 맡는다 → ViewModel 은 Excel 을 몰라도 된다.
/// </summary>
public interface IExcelBridge
{
    /// <summary>지금 Excel 에서 선택된 것 (칸 또는 그림). Excel 이 없으면 null.</summary>
    ExcelSelection? GetSelection();

    /// <summary>현재 시트에서 그림·글상자들(시트 안 순서, 0부터)의 위치를 읽는다.</summary>
    List<ShapePlacement> ReadPlacements(IEnumerable<int> shapeIndexes);

    /// <summary>그림·글상자들을 저장된 위치로 옮겨 보여 준다 (미리보기만, 파일 저장 안 함).</summary>
    void ApplyPlacements(IEnumerable<ShapePlacement> placements);

    /// <summary>항목을 누르면 Excel 에서 해당 칸/그림을 선택해 보여 준다.</summary>
    void Reveal(string? cell, IEnumerable<int>? shapeIndexes);

    /// <summary>지금 시트의 (행, 열) 칸에 보이는 글자 (머리글 자동 인식용). 1부터.</summary>
    string GetCellText(int row, int col);
}

/// <summary>Excel 선택 상태</summary>
/// <param name="IsShapes">그림·글상자가 선택됨 (아니면 칸)</param>
/// <param name="ActiveCell">칸: 현재 칸 주소(예: "D12") / 그림: 그림 왼쪽 위 칸</param>
/// <param name="Row">선택 범위 첫 행 (1부터)</param>
/// <param name="Col">선택 범위 첫 열 (1부터)</param>
/// <param name="RowCount">선택 범위 행 수</param>
/// <param name="ShapeNames">선택된 그림·글상자 이름들</param>
/// <param name="ShapeIndexes">선택된 그림·글상자의 시트 안 순서 (0부터)</param>
/// <param name="ColCount">선택 범위 열 수 (표 머리글 자동 인식에 쓴다)</param>
public record ExcelSelection(bool IsShapes, string ActiveCell, int Row, int Col, int RowCount,
                             List<string> ShapeNames, List<int> ShapeIndexes, int ColCount = 1)
{
    /// <summary>선택이 바뀌었는지 비교하는 용도</summary>
    public string Signature => IsShapes ? "S:" + string.Join(",", ShapeNames) : $"R:{ActiveCell}:{Row}:{Col}:{RowCount}";
}

/// <summary>Excel 에서 무엇을 고르는 중인지</summary>
public enum PickKind
{
    Cell,       // 값을 쓸 칸 하나
    ValueCell,  // 사진 옆 측정값 칸
    Picture,    // 사진 칸(그림)
    Shapes,     // 옮길 그림·글상자 (여러 개, 확인 버튼)
    Rows,       // 표의 데이터 행 범위 (드래그 후 확인)
    Lines,      // 여러 줄 텍스트 칸들 (드래그 후 확인)
    Column,     // 표의 열 하나 (그 열의 아무 칸 클릭)
    TotalCell,  // 숫자 누계: 누계 칸
    PrevCell,   // 숫자 누계: 전일 칸 (없으면 건너뛰기)
    PhotoRange,    // 사진 구역: 사진 칸 자리 (드래그 또는 그림 클릭 후 확인)
    PhotoCaption,  // 사진 구역: 방금 지정한 칸의 설명 칸 (없으면 건너뛰기)
    SingleRange,   // 사진 구역: 1장일 때 자리 (없으면 건너뛰기)
    SingleCaption, // 사진 구역: 1장일 때 설명 칸 (없으면 건너뛰기)
}

// ════════════════════════════════════════════════════════════════════
// 항목 종류 (추가할 때 고르는 목록)
// ════════════════════════════════════════════════════════════════════

public record ItemTypeOption(string Icon, string Label, string Type, string? FixedId = null, string Hint = "")
{
    public string Display => $"{Icon}  {Label}";
    public override string ToString() => Display;

    public bool AsksSuffix => Type is FieldTypes.Input or FieldTypes.Photo or BlockTypes.Value;
    public bool AsksTime   => Type == FieldTypes.Photo;
    public bool HasHint    => Hint.Length > 0;

    /// <summary>[+ 항목 추가] 에 보이는 종류. 뜻(날씨·장비 등)은 이름과 규칙으로 사용자가 정한다.</summary>
    public static readonly IReadOnlyList<ItemTypeOption> All =
    [
        new("🔤", "값 — 글자·숫자 하나",                BlockTypes.Value,   Hint: "날씨, 공종, 측정값처럼 칸 하나. 매일 입력 / 어제 값 그대로 / 고정값 중에서 고릅니다."),
        new("📅", "날짜 — 형식 고르기",                  BlockTypes.Date,    Hint: "고른 날짜를 원하는 모양(2026.11.18 / 2026년 11월 18일 수요일 …)으로 씁니다."),
        new("📝", "여러 줄 글 — 작업사항 등",            BlockTypes.Lines,   Hint: "칸 여러 개에 한 줄씩. '어제 명일 → 오늘 금일' 처럼 처음 값을 정할 수 있습니다."),
        new("➕", "숫자 누계 — 오늘 + 누계",             BlockTypes.Counter, Hint: "오늘 숫자를 입력하면 누계 = 어제 누계 + 오늘. 폐기물 반출 대수, 철거동 수 등."),
        new("📊", "표 — 줄마다 전일·금일·누계",          BlockTypes.Table,   Hint: "장비·인원 투입현황처럼 이름 줄마다 숫자. 열마다 역할을 클릭으로 정합니다."),
        new("👥", "명단 — 오늘 나온 사람",               BlockTypes.Roster,  Hint: "출력일보처럼 한 줄에 한 사람(직종·성명·작업내용). 사람 목록에서 체크합니다."),
        new("🧮", "계산 — 다른 항목 합계·이름 나열",     BlockTypes.Calc,    Hint: "TBM 인원 = 인원 표 합계 + 수장재, 투입 장비 이름 나열 등."),
        new("📷", "사진 구역 — 그날 장수만큼 채우기",    BlockTypes.Photos,  Hint: "사진대지처럼 사진 칸 여러 개(+설명 칸). 사진이 1장인 날은 따로 정한 자리(두 칸 합친 자리 등)에 넣을 수 있습니다."),
        new("📷", "사진 (+ 옆 칸 측정값)",               FieldTypes.Photo),
        new("📍", "그림·글상자 위치 (장비·측정기 표시)", FieldTypes.ShapePosition),
        new("✏️", "시트마다 따로 입력하는 값",            FieldTypes.Input,   Hint: "같은 모양 시트 여러 개에 시트마다 다른 값을 넣을 때 (예: 측정 장소별 값)."),
    ];

    /// <summary>예전 방식 종류 (불러온 설정에만 보이고, 새로 추가하는 목록에는 없다)</summary>
    private static readonly IReadOnlyList<ItemTypeOption> Legacy =
    [
        new("📅", "날짜 — 2026.11.05 (예전 방식)",              FieldTypes.DateSimple),
        new("📅", "날짜 — 2026년 11월 05일 목요일 (예전 방식)", FieldTypes.DateKorean),
        new("🌤", "날씨 (예전 방식)",                           FieldTypes.Text, "weather"),
        new("🔤", "글자 (예전 방식)",                           FieldTypes.Text),
        new("📝", "금일 작업사항 (예전 방식)",                  FieldTypes.Multiline, "today_work"),
        new("📝", "명일 작업사항 (예전 방식)",                  FieldTypes.Multiline, "tomorrow_work"),
        new("📝", "여러 줄 (예전 방식)",                        FieldTypes.Multiline),
        new("🚜", "장비 투입현황 표 (예전 방식)",               FieldTypes.EquipmentTable),
        new("👷", "인원 투입현황 표 (예전 방식)",               FieldTypes.PersonnelTable),
        new("🚜", "투입 장비명 (예전 방식)",                    FieldTypes.ActiveEquipmentNames),
        new("👷", "TBM 인원 (예전 방식)",                       FieldTypes.TbmPersonnel),
        new("✅", "참석율 100 (예전 방식)",                     FieldTypes.AttendanceRate),
    ];

    public static ItemTypeOption For(string type, string id) =>
        All.Concat(Legacy).FirstOrDefault(o => o.Type == type && (o.FixedId == null || o.FixedId == id))
        ?? All.Concat(Legacy).FirstOrDefault(o => o.Type == type)
        ?? new("•", type, type);
}

/// <summary>콤보 상자의 선택지 하나 (화면 글자 + 저장 값)</summary>
public record Choice(string Label, string Value)
{
    public override string ToString() => Label;
}

// ════════════════════════════════════════════════════════════════════
// 항목 카드 하나
// ════════════════════════════════════════════════════════════════════

/// <summary>설정 화면 오른쪽 목록의 항목 카드 하나 (= config 의 field 하나)</summary>
public class SetupItemViewModel : ViewModelBase
{
    private readonly ISetupHost _host;

    public string Id   { get; }
    public string Type { get; }
    public ItemTypeOption Option { get; }
    public string Icon => Option.Icon;
    public string TypeLabel => Option.Label;

    private string _label;
    public string Label { get => _label; set { if (SetField(ref _label, value)) OnPropertyChanged(nameof(KeyHint)); } }

    // ── 지정 정보 ────────────────────────────────────────────────────
    private string? _cell, _valueCell, _suffix, _time, _placeholder, _totalCell, _prevCell, _sheet;
    public string? Cell        { get => _cell;        set { if (SetField(ref _cell, value)) Changed(); } }
    public string? ValueCell   { get => _valueCell;   set { if (SetField(ref _valueCell, value)) Changed(); } }
    public string? Suffix      { get => _suffix;      set { if (SetField(ref _suffix, value)) Changed(); } }
    public string? Time        { get => _time;        set { if (SetField(ref _time, value)) Changed(); } }
    public string? Placeholder { get => _placeholder; set { if (SetField(ref _placeholder, value)) Changed(); } }
    public string? TotalCell   { get => _totalCell;   set { if (SetField(ref _totalCell, value)) Changed(); } }
    public string? PrevCell    { get => _prevCell;    set { if (SetField(ref _prevCell, value)) Changed(); } }

    /// <summary>"적용할 시트" 가 아닌 시트에서 지정했으면 그 시트 이름 (그 시트에만 쓴다)</summary>
    public string? Sheet { get => _sheet; set { if (SetField(ref _sheet, value)) Changed(); } }
    public bool HasOwnSheet => !string.IsNullOrWhiteSpace(_sheet);

    public List<string> Cells { get; private set; } = [];
    public List<string> ShapeNames { get; private set; } = [];
    public List<int> ShapeIndexes { get; private set; } = [];
    public ObservableCollection<PresetItemViewModel> Presets { get; } = [];

    // 표·명단
    public int StartRow { get; private set; }
    public int EndRow   { get; private set; }
    public Dictionary<string, string> Cols { get; } = [];   // 역할 → 열 문자

    // 사진 구역
    public List<PhotoArea> PhotoAreas { get; private set; } = [];
    public PhotoArea? SinglePhotoArea { get; private set; }

    /// <summary>사진 구역 카드에 보여 줄 칸 목록 (한 줄에 한 칸)</summary>
    public string PhotoAreasText =>
        PhotoAreas.Count == 0 ? "아직 칸이 없습니다. [＋ 사진 칸 추가]를 누르세요."
        : string.Join("\n", PhotoAreas.Select((a, i) => $"칸 {i + 1}: {AreaText(a)}"))
          + "\n" + (SinglePhotoArea == null ? "1장일 때: 칸 1 에 넣음" : $"1장일 때: {AreaText(SinglePhotoArea)}");

    private static string AreaText(PhotoArea a) =>
        a.Range + (string.IsNullOrEmpty(a.CaptionCell) ? " (설명 없음)" : $" · 설명 {a.CaptionCell}");

    private string _newPresetName = "";
    public string NewPresetName { get => _newPresetName; set => SetField(ref _newPresetName, value); }

    // ── 블록 규칙 ────────────────────────────────────────────────────

    private string _key = "";
    /// <summary>값 이름. 다른 문서와 같은 이름이면 메인 화면에서 한 번만 입력한다.</summary>
    public string Key { get => _key; set { if (SetField(ref _key, value)) OnPropertyChanged(nameof(KeyHint)); } }
    public string KeyHint => string.IsNullOrWhiteSpace(_key) ? $"(비우면 '{Label}')" : "";

    public static IReadOnlyList<Choice> ValueSourceList { get; } =
    [
        new("매일 입력", BlockTypes.SourceInput), new("어제 값 그대로 (안 바꿈)", BlockTypes.SourceKeep),
        new("항상 같은 값", BlockTypes.SourceFixed),
    ];
    // 화면 바인딩용 (WPF 바인딩은 인스턴스 속성만 읽는다)
    public IReadOnlyList<Choice> ValueSources => ValueSourceList;
    public IReadOnlyList<Choice> CalcKinds => CalcKindList;
    /// <summary>다른 항목 이름 목록 (명단 이름 고르기 등)</summary>
    public List<string> OtherKeys => _host.KnownKeys().Where(k => k != EffectiveKey).ToList();

    private Choice _valueSource = ValueSourceList[0];
    public Choice ValueSource
    {
        get => _valueSource;
        set { if (SetField(ref _valueSource, value)) { OnPropertyChanged(nameof(IsFixedValue)); Changed(); } }
    }
    public bool IsFixedValue => _valueSource.Value == BlockTypes.SourceFixed;

    private string _fixedValue = "";
    public string FixedValue { get => _fixedValue; set => SetField(ref _fixedValue, value); }

    public IReadOnlyList<string> DateFormats => BlockTypes.DateFormats;
    private string _dateFormat = BlockTypes.DateFormats[0];
    public string DateFormat
    {
        get => _dateFormat;
        set { if (SetField(ref _dateFormat, value)) { OnPropertyChanged(nameof(DatePreview)); Changed(); } }
    }
    /// <summary>오늘 날짜로 미리 보기 (예: "2026.11.18")</summary>
    public string DatePreview => "→ " + BlockTypes.FormatDate(_dateFormat, DateTime.Today);

    private string _prefix = "";
    public string Prefix { get => _prefix; set => SetField(ref _prefix, value); }

    private string _startFrom = "";
    /// <summary>여러 줄의 처음 값: "" 빈 칸 / "same" 어제 이 항목 / 다른 항목 이름</summary>
    public string StartFrom { get => _startFrom; set => SetField(ref _startFrom, value); }
    /// <summary>처음 값 고르기 목록 (다른 항목 이름은 host 가 알려 준다)</summary>
    public List<Choice> StartFromChoices =>
        [new("빈 칸에서 시작", ""), new("어제 이 항목 그대로", "same"),
         .. _host.KnownKeys().Where(k => k != EffectiveKey).Select(k => new Choice($"어제 '{k}' 에서 (예: 금일 ← 명일)", k))];

    public static IReadOnlyList<Choice> CalcKindList { get; } =
    [
        new("합계 (표의 금일·숫자·명단 인원 더하기)", BlockTypes.CalcSum),
        new("이름 나열 (표에서 금일이 있는 줄)", BlockTypes.CalcNames),
        new("명단 인원 수", BlockTypes.CalcCount),
    ];
    private Choice _calcKind = CalcKindList[0];
    public Choice CalcKind { get => _calcKind; set { if (SetField(ref _calcKind, value)) Changed(); } }

    private string _sourcesText = "";
    /// <summary>계산에 쓸 항목 이름들 (쉼표로)</summary>
    public string SourcesText { get => _sourcesText; set { if (SetField(ref _sourcesText, value)) { Changed(); RebuildSourceChoices(); } } }

    /// <summary>계산에 쓸 항목: 이름을 타이핑하지 않고 체크해서 고른다 (오타 방지)</summary>
    public ObservableCollection<KeyChoice> SourceChoices { get; } = [];

    private List<string> SourceList() =>
        _sourcesText.Split(',', StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Where(x => x.Length > 0).ToList();

    private void RebuildSourceChoices()
    {
        if (!IsCalc) return;
        var current = SourceList();
        var known = OtherKeys;
        SourceChoices.Clear();
        foreach (var k in known) SourceChoices.Add(new KeyChoice(k, current.Contains(k), missing: false, OnSourceToggled));
        foreach (var k in current.Where(c => !known.Contains(c)))   // 지금은 없는 이름 (오타·지운 항목)
            SourceChoices.Add(new KeyChoice(k, true, missing: true, OnSourceToggled));
    }

    private void OnSourceToggled()
    {
        _sourcesText = string.Join(", ", SourceChoices.Where(c => c.IsChecked).Select(c => c.Key));
        OnPropertyChanged(nameof(SourcesText));
        Changed();
    }

    // ── 머리글 자동 인식 안내 ────────────────────────────────────────
    private string _autoNote = "";
    /// <summary>표·명단 행을 드래그했을 때 머리글에서 찾은 열 (예: "머리글에서 찾음: 이름 A · 금일 H")</summary>
    public string AutoNote { get => _autoNote; set { if (SetField(ref _autoNote, value)) OnPropertyChanged(nameof(HasAutoNote)); } }
    public bool HasAutoNote => _autoNote.Length > 0;

    // ── 짝짓기 표 (표의 줄 ↔ 명단 직종) ─────────────────────────────
    public ObservableCollection<RoleMapRow> RoleMapRows { get; } = [];
    public bool HasRoleMapRows => RoleMapRows.Count > 0;
    public bool HasRosterKey => !string.IsNullOrWhiteSpace(_rosterKey);
    public ICommand LoadRoleMapCommand { get; private set; } = null!;

    /// <summary>
    /// 짝짓기 표를 만든다. 줄마다 직종 체크박스. 이미 정한 짝짓기가 있으면 그대로, 없으면 줄 이름과 같은 직종만 체크.
    /// </summary>
    public void SetRoleMapData(IEnumerable<string> rowNames, IEnumerable<string> jobs)
    {
        var map = ParseRoleMap(_roleMapText) ?? [];
        var jobList = jobs.Select(j => j.Trim()).Where(j => j.Length > 0).Distinct().ToList();
        RoleMapRows.Clear();
        foreach (var row in rowNames.Select(r => r.Trim()).Where(r => r.Length > 0).Distinct())
        {
            var picked = map.TryGetValue(row, out var m) ? m : jobList.Where(j => j == row).ToList();
            var all = jobList.Union(picked).ToList();
            RoleMapRows.Add(new RoleMapRow(row, all.Select(j => new KeyChoice(j, picked.Contains(j), false, OnRoleMapToggled))));
        }
        OnPropertyChanged(nameof(HasRoleMapRows));
    }

    private void OnRoleMapToggled()
    {
        // 기본(줄 이름과 같은 직종 하나)과 다른 줄만 저장해 글자를 짧게
        var map = new Dictionary<string, List<string>>();
        foreach (var r in RoleMapRows)
        {
            var picked = r.Jobs.Where(j => j.IsChecked).Select(j => j.Key).ToList();
            if (picked.Count > 0 && !(picked.Count == 1 && picked[0] == r.RowName)) map[r.RowName] = picked;
        }
        _roleMapText = FormatRoleMap(map.Count == 0 ? null : map);
        OnPropertyChanged(nameof(RoleMapText));
    }
    public string KnownKeysText => "쓸 수 있는 이름: " + string.Join(", ", _host.KnownKeys().Where(k => k != EffectiveKey));

    private string _rosterKey = "";
    /// <summary>표: 금일을 셀 명단 이름 (비우면 화면에서 입력)</summary>
    public string RosterKey
    {
        get => _rosterKey;
        set
        {
            if (!SetField(ref _rosterKey, value)) return;
            OnPropertyChanged(nameof(HasRosterKey));
            // 아는 명단 이름을 고르면 짝짓기 표를 바로 채운다
            if (StartRow > 0 && _host.KnownKeys().Contains(value.Trim())) _host.LoadRoleMap(this);
        }
    }

    private string _roleMapText = "";
    /// <summary>표: 줄 이름 ↔ 명단 직종 짝짓기. 예) "관리자 = 소장, 차장, 대리 / 장비 = 장비"</summary>
    public string RoleMapText { get => _roleMapText; set => SetField(ref _roleMapText, value); }

    public string EffectiveKey => string.IsNullOrWhiteSpace(_key) ? Label.Trim() : _key.Trim();

    // ── 화면 표시용 ──────────────────────────────────────────────────
    public bool IsHwpx        => _host.IsHwpx;
    public bool IsExcel       => !_host.IsHwpx;
    public bool IsPhoto       => Type == FieldTypes.Photo;
    public bool IsShape       => Type == FieldTypes.ShapePosition && !IsHwpx;
    public bool IsEquipTable  => Type == FieldTypes.EquipmentTable && !IsHwpx;
    public bool IsPersonTable => Type == FieldTypes.PersonnelTable && !IsHwpx;
    public bool HasSuffix     => Type is FieldTypes.Input or FieldTypes.Photo || (Type == BlockTypes.Value && !IsFixedValue);

    public bool IsBlock   => BlockTypes.IsBlock(Type);
    public bool HasKey    => IsBlock && Type != BlockTypes.Date;
    public bool IsValue   => Type == BlockTypes.Value;
    public bool IsDate    => Type == BlockTypes.Date;
    public bool IsLines   => Type == BlockTypes.Lines;
    public bool IsCounter => Type == BlockTypes.Counter && !IsHwpx;
    public bool IsTable   => Type == BlockTypes.Table;
    public bool IsCalc    => Type == BlockTypes.Calc;
    public bool IsPhotos  => Type == BlockTypes.Photos && !IsHwpx;
    public bool IsLegacy  => BlockMigration.IsLegacy(Type);

    /// <summary>카드에 보여 줄 지정 내용 요약 (예: "사진 D12 · 값 E12 · 10:00")</summary>
    public string Summary => IsHwpx ? "" : (HasOwnSheet ? $"[{Sheet}] " : "") + Type switch
    {
        FieldTypes.Photo => Join(Cell is null ? null : $"사진 {Cell}", ValueCell is null ? null : $"값 {ValueCell}",
                                 string.IsNullOrEmpty(Time) ? null : Time),
        FieldTypes.ShapePosition => ShapeNames.Count == 0 ? "" : $"{string.Join(", ", ShapeNames)} · 저장된 위치 {Presets.Count}개",
        FieldTypes.Multiline or BlockTypes.Lines => Cells.Count == 0 ? "" : $"{Cells[0]} ~ {Cells[^1]} ({Cells.Count}줄)",
        BlockTypes.Counter => Join(Cell is null ? null : $"오늘 {Cell}", TotalCell is null ? null : $"누계 {TotalCell}",
                                   PrevCell is null ? null : $"전일 {PrevCell}"),
        FieldTypes.EquipmentTable or FieldTypes.PersonnelTable or BlockTypes.Table or BlockTypes.Roster =>
            StartRow == 0 ? "" : $"{StartRow}~{EndRow}행 · " + string.Join(" ", ColKeys.Where(k => ColOf(k) != null).Select(k => $"{ColTitle(k)} {ColOf(k)}")),
        BlockTypes.Calc => SourcesText.Length == 0 ? Cell ?? "" : Join(Cell, $"= {CalcKind.Label.Split(' ')[0]}({SourcesText})"),
        BlockTypes.Photos => PhotoAreas.Count == 0 ? "" : $"사진 칸 {PhotoAreas.Count}개" + (SinglePhotoArea == null ? "" : " · 1장일 때 자리 있음"),
        _ => Cell ?? (IsValue ? "칸 없음 (계산에만 사용)" : ""),
    };

    /// <summary>엑셀에서 지정이 끝났는지 (카드의 ✓ / "지정 필요")</summary>
    public bool IsComplete => IsHwpx ? !string.IsNullOrWhiteSpace(Placeholder) : Type switch
    {
        FieldTypes.Photo         => Cell != null,
        FieldTypes.ShapePosition => ShapeNames.Count > 0,
        FieldTypes.Multiline or BlockTypes.Lines => Cells.Count > 0,
        FieldTypes.EquipmentTable or FieldTypes.PersonnelTable => StartRow > 0 && ColKeys.All(k => ColOf(k) != null),
        BlockTypes.Table or BlockTypes.Roster => StartRow > 0 && RequiredColKeys.All(k => ColOf(k) != null),
        BlockTypes.Counter => Cell != null || TotalCell != null,
        BlockTypes.Value   => true,   // 칸이 없으면 계산에만 쓰는 입력
        BlockTypes.Calc    => Cell != null && SourcesText.Trim().Length > 0,
        BlockTypes.Photos  => PhotoAreas.Count > 0,
        _ => Cell != null,
    };

    /// <summary>표·명단에서 지정할 열 (순서대로 안내)</summary>
    public string[] ColKeys => Type switch
    {
        FieldTypes.EquipmentTable => ["name", "prev", "today"],
        FieldTypes.PersonnelTable => ["role", "prevCum", "today"],
        BlockTypes.Table  => BlockTypes.TableRoles.Select(r => r.Role).ToArray(),
        BlockTypes.Roster => BlockTypes.RosterRoles.Select(r => r.Role).ToArray(),
        _ => [],
    };

    /// <summary>꼭 지정해야 하는 열 (나머지는 [건너뛰기] 가능)</summary>
    public string[] RequiredColKeys => Type switch
    {
        BlockTypes.Table  => BlockTypes.TableRoles.Where(r => r.Required).Select(r => r.Role).ToArray(),
        BlockTypes.Roster => BlockTypes.RosterRoles.Where(r => r.Required).Select(r => r.Role).ToArray(),
        _ => ColKeys,
    };

    public bool IsOptionalCol(string key) => !RequiredColKeys.Contains(key);

    public static string ColTitle(string key) => key switch
    {
        "role" => "직종", "prevCum" => "전일누계",
        _ => BlockTypes.RoleTitle(key),   // 이름·규격·전일·금일·누계·직종·성명·작업내용
    };

    public string? ColOf(string key) => Cols.TryGetValue(key, out var v) ? v : null;

    /// <summary>표 카드의 열 버튼들 (예: "이름 B", "전일 ?")</summary>
    public List<ColumnButton> ColumnButtons =>
        ColKeys.Select(k => new ColumnButton(k, $"{ColTitle(k)} {ColOf(k) ?? (IsOptionalCol(k) ? "-" : "?")}")).ToList();

    // ── 커맨드 ───────────────────────────────────────────────────────
    public ICommand PickCommand       { get; }
    public ICommand PickValueCommand  { get; }
    public ICommand PickColumnCommand { get; }
    public ICommand PickTotalCommand  { get; }
    public ICommand PickPrevCommand   { get; }
    public ICommand RemoveCommand     { get; }
    public ICommand RevealCommand     { get; }
    public ICommand SavePresetCommand { get; }
    public ICommand AddPhotoAreaCommand    { get; }
    public ICommand PickSinglePhotoCommand { get; }
    public ICommand ClearPhotoAreasCommand { get; }

    public SetupItemViewModel(ISetupHost host, string id, string label, string type)
    {
        _host  = host;
        Id     = id;
        _label = label;
        Type   = type;
        Option = ItemTypeOption.For(type, id);

        PickCommand       = new RelayCommand(_ =>
        {
            // 사진 구역: [Excel 에서 지정] 은 칸을 처음부터 다시 지정 (하나 더 넣기는 [＋ 사진 칸 추가])
            if (Type == BlockTypes.Photos) { PhotoAreas = []; SinglePhotoArea = null; Changed(); }
            _host.StartPick(this, MainPickKind);
        });
        PickValueCommand  = new RelayCommand(_ => _host.StartPick(this, PickKind.ValueCell));
        PickColumnCommand = new RelayCommand(p => _host.StartPick(this, PickKind.Column, p as string));
        PickTotalCommand  = new RelayCommand(_ => _host.StartPick(this, PickKind.TotalCell));
        PickPrevCommand   = new RelayCommand(_ => _host.StartPick(this, PickKind.PrevCell));
        RemoveCommand     = new RelayCommand(_ => _host.RemoveItem(this));
        RevealCommand     = new RelayCommand(_ => _host.RevealItem(this));
        SavePresetCommand = new RelayCommand(_ => _host.SavePreset(this));
        LoadRoleMapCommand = new RelayCommand(_ => _host.LoadRoleMap(this));
        AddPhotoAreaCommand    = new RelayCommand(_ => _host.StartPick(this, PickKind.PhotoRange));
        PickSinglePhotoCommand = new RelayCommand(_ => _host.StartPick(this, PickKind.SingleRange), _ => PhotoAreas.Count > 0);
        ClearPhotoAreasCommand = new RelayCommand(_ => { PhotoAreas = []; SinglePhotoArea = null; Changed(); }, _ => PhotoAreas.Count > 0);
    }

    /// <summary>이 항목을 처음 지정할 때 Excel 에서 무엇을 골라야 하는지</summary>
    public PickKind MainPickKind => Type switch
    {
        FieldTypes.Photo         => PickKind.Picture,
        FieldTypes.ShapePosition => PickKind.Shapes,
        FieldTypes.Multiline or BlockTypes.Lines => PickKind.Lines,
        FieldTypes.EquipmentTable or FieldTypes.PersonnelTable or BlockTypes.Table or BlockTypes.Roster => PickKind.Rows,
        BlockTypes.Photos => PickKind.PhotoRange,
        _ => PickKind.Cell,
    };

    // ── 사진 구역 지정 ───────────────────────────────────────────────
    public void AddPhotoArea(string range) { PhotoAreas = [.. PhotoAreas, new PhotoArea { Range = range }]; Changed(); }
    /// <summary>방금 추가한 칸의 설명 칸</summary>
    public void SetLastCaption(string cell)
    {
        if (PhotoAreas.Count == 0) return;
        PhotoAreas = [.. PhotoAreas.Take(PhotoAreas.Count - 1), new PhotoArea { Range = PhotoAreas[^1].Range, CaptionCell = cell }];
        Changed();
    }
    public void SetSingleArea(string range) { SinglePhotoArea = new PhotoArea { Range = range }; Changed(); }
    public void SetSingleCaption(string cell)
    {
        if (SinglePhotoArea == null) return;
        SinglePhotoArea = new PhotoArea { Range = SinglePhotoArea.Range, CaptionCell = cell };
        Changed();
    }

    // ── 값 설정 (ViewModel 에서 호출) ────────────────────────────────

    public void SetShapes(List<string> names, List<int> indexes) { ShapeNames = names; ShapeIndexes = indexes; Changed(); }
    public void SetLines(List<string> cells)  { Cells = cells; Changed(); }
    public void SetRows(int start, int end)   { StartRow = start; EndRow = end; Changed(); }
    public void SetColumn(string key, string col) { Cols[key] = col; Changed(); }

    /// <summary>숫자 누계: 전일 칸이 없다고 [건너뛰기] 했는지 (다시 묻지 않게)</summary>
    public bool PrevSkipped { get; private set; }
    public void SetPrevSkipped() => PrevSkipped = true;

    public void AddPreset(ShapePreset p)
    {
        Presets.Add(new PresetItemViewModel(p, this, _host));
        Changed();
    }

    public void RemovePreset(PresetItemViewModel p) { Presets.Remove(p); Changed(); }

    public void Changed()
    {
        OnPropertyChanged(nameof(Summary));
        OnPropertyChanged(nameof(IsComplete));
        OnPropertyChanged(nameof(ColumnButtons));
        OnPropertyChanged(nameof(HasSuffix));
        OnPropertyChanged(nameof(HasOwnSheet));
        OnPropertyChanged(nameof(PhotoAreasText));
    }

    /// <summary>엑셀/한글 전환 시 표시 갱신</summary>
    public void RefreshMode()
    {
        foreach (var n in new[] { nameof(IsHwpx), nameof(IsExcel), nameof(IsShape), nameof(IsEquipTable),
                                  nameof(IsPersonTable), nameof(IsCounter), nameof(IsPhotos), nameof(Summary), nameof(IsComplete) })
            OnPropertyChanged(n);
    }

    /// <summary>다른 항목이 추가·이름이 바뀌었을 때 (처음 값·계산 목록 갱신)</summary>
    public void RefreshKnownKeys()
    {
        OnPropertyChanged(nameof(StartFromChoices));
        OnPropertyChanged(nameof(KnownKeysText));
        OnPropertyChanged(nameof(OtherKeys));
        RebuildSourceChoices();
    }

    private static string Join(params string?[] parts) => string.Join(" · ", parts.Where(p => !string.IsNullOrEmpty(p)));

    // ── 짝짓기 글자 ↔ 사전 ───────────────────────────────────────────

    /// <summary>"관리자 = 소장, 차장, 대리 / 장비 = 장비" → 사전</summary>
    public static Dictionary<string, List<string>>? ParseRoleMap(string text)
    {
        var map = new Dictionary<string, List<string>>();
        foreach (var part in text.Split(new[] { '/', ';', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            int eq = part.IndexOf('=');
            if (eq <= 0) continue;
            string row = part[..eq].Trim();
            var jobs = part[(eq + 1)..].Split(',', StringSplitOptions.RemoveEmptyEntries).Select(j => j.Trim()).Where(j => j.Length > 0).ToList();
            if (row.Length > 0 && jobs.Count > 0) map[row] = jobs;
        }
        return map.Count == 0 ? null : map;
    }

    public static string FormatRoleMap(Dictionary<string, List<string>>? map) =>
        map == null ? "" : string.Join(" / ", map.Select(kv => $"{kv.Key} = {string.Join(", ", kv.Value)}"));

    // ── config 변환 ──────────────────────────────────────────────────

    public static SetupItemViewModel FromConfig(ISetupHost host, FieldConfig f)
    {
        var item = new SetupItemViewModel(host, f.Id, f.Label, f.Type)
        {
            _cell = f.Cell, _valueCell = f.ValueCell, _suffix = f.Suffix, _time = f.Time,
            _placeholder = f.Placeholder, _totalCell = f.TotalCell, _prevCell = f.PrevCell, _sheet = f.Sheet,
            _key = f.Key ?? "", _fixedValue = f.FixedValue ?? "", _prefix = f.Prefix ?? "", _startFrom = f.StartFrom ?? "",
            _sourcesText = string.Join(", ", f.Sources ?? []), _rosterKey = f.RosterKey ?? "",
            _roleMapText = FormatRoleMap(f.RoleMap),
        };
        item._valueSource = ValueSourceList.FirstOrDefault(c => c.Value == (f.ValueSource ?? BlockTypes.SourceInput)) ?? ValueSourceList[0];
        item._calcKind    = CalcKindList.FirstOrDefault(c => c.Value == (f.CalcKind ?? BlockTypes.CalcSum)) ?? CalcKindList[0];
        if (!string.IsNullOrWhiteSpace(f.DateFormat)) item._dateFormat = f.DateFormat;
        item.Cells      = f.Cells ?? [];
        item.ShapeNames = f.ShapeNames ?? [];
        item.ShapeIndexes = f.ShapeIndexes ?? [];
        item.StartRow   = f.StartRow;
        item.EndRow     = f.EndRow;
        item.PhotoAreas = (f.PhotoAreas ?? []).Select(a => new PhotoArea { Range = a.Range, CaptionCell = a.CaptionCell }).ToList();
        item.SinglePhotoArea = f.SinglePhotoArea is { } s ? new PhotoArea { Range = s.Range, CaptionCell = s.CaptionCell } : null;
        void Col(string k, string? v) { if (!string.IsNullOrEmpty(v)) item.Cols[k] = v; }
        Col("name", f.NameCol); Col("prev", f.PrevDayCol); Col("today", f.TodayCol);
        item._specCol = f.SpecCol;   // 화면에서 고르지 않는 값 — 그대로 보존
        Col("role", f.RoleCol); Col("prevCum", f.PrevCumulativeCol);
        foreach (var c in f.Columns ?? []) Col(c.Role, c.Col);
        foreach (var p in f.Presets ?? []) item.Presets.Add(new PresetItemViewModel(p, item, host));
        return item;
    }

    private string? _specCol;

    public FieldConfig ToConfig()
    {
        var f = new FieldConfig { Id = Id, Label = Label.Trim(), Type = Type };
        f.Sheet = NullIfEmpty(Sheet);
        if (IsBlock)
        {
            f.Key = NullIfEmpty(Key);
            switch (Type)
            {
                case BlockTypes.Value:
                    f.ValueSource = ValueSource.Value == BlockTypes.SourceInput ? null : ValueSource.Value;
                    f.FixedValue  = IsFixedValue ? FixedValue : null;
                    f.Suffix      = IsFixedValue ? null : NullIfEmpty(Suffix);
                    break;
                case BlockTypes.Date:
                    f.DateFormat = DateFormat;
                    break;
                case BlockTypes.Lines:
                    f.Prefix    = string.IsNullOrEmpty(Prefix) ? null : Prefix;
                    f.StartFrom = NullIfEmpty(StartFrom);
                    break;
                case BlockTypes.Calc:
                    f.CalcKind = CalcKind.Value;
                    f.Sources  = SourcesText.Split(',', StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim())
                                            .Where(x => x.Length > 0).ToList();
                    break;
                case BlockTypes.Table:
                    f.RosterKey = NullIfEmpty(RosterKey);
                    f.RoleMap   = f.RosterKey == null ? null : ParseRoleMap(RoleMapText);
                    break;
            }
        }

        if (IsHwpx)
        {
            f.Placeholder = string.IsNullOrWhiteSpace(Placeholder) ? null : Placeholder.Trim();
            if (!IsBlock) f.Suffix = NullIfEmpty(Suffix);
            return f;
        }

        switch (Type)
        {
            case FieldTypes.Photo:
                f.Cell = Cell; f.ValueCell = ValueCell; f.Suffix = NullIfEmpty(Suffix); f.Time = NullIfEmpty(Time);
                break;
            case FieldTypes.ShapePosition:
                f.ShapeNames   = [.. ShapeNames];
                f.ShapeIndexes = [.. ShapeIndexes];
                f.Presets      = Presets.Select(p => p.Preset).ToList();
                break;
            case FieldTypes.Multiline:
            case BlockTypes.Lines:
                f.Cells = [.. Cells];
                break;
            case FieldTypes.EquipmentTable:
                f.StartRow = StartRow; f.EndRow = EndRow;
                f.NameCol = ColOf("name"); f.PrevDayCol = ColOf("prev"); f.TodayCol = ColOf("today");
                f.SpecCol = _specCol;
                break;
            case FieldTypes.PersonnelTable:
                f.StartRow = StartRow; f.EndRow = EndRow;
                f.RoleCol = ColOf("role"); f.PrevCumulativeCol = ColOf("prevCum"); f.TodayCol = ColOf("today");
                break;
            case BlockTypes.Table:
            case BlockTypes.Roster:
                f.StartRow = StartRow; f.EndRow = EndRow;
                f.Columns = ColKeys.Where(k => ColOf(k) != null).Select(k => new ColumnRole { Role = k, Col = ColOf(k)! }).ToList();
                break;
            case BlockTypes.Counter:
                f.Cell = Cell; f.TotalCell = TotalCell; f.PrevCell = PrevCell;
                break;
            case BlockTypes.Photos:
                f.PhotoAreas      = PhotoAreas.Select(a => new PhotoArea { Range = a.Range, CaptionCell = a.CaptionCell }).ToList();
                f.SinglePhotoArea = SinglePhotoArea is { } s ? new PhotoArea { Range = s.Range, CaptionCell = s.CaptionCell } : null;
                break;
            case BlockTypes.Value:
            case BlockTypes.Date:
            case BlockTypes.Calc:
                f.Cell = Cell;
                break;
            default:
                f.Cell = Cell; f.Suffix = NullIfEmpty(Suffix);
                break;
        }
        return f;
    }

    private static string? NullIfEmpty(string? s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();
}

/// <summary>저장된 위치 하나 (누르면 Excel 에서 그 위치로 옮겨 보여 줌)</summary>
public class PresetItemViewModel
{
    public ShapePreset Preset { get; }
    public string Name => Preset.Name;
    public ICommand ShowCommand   { get; }
    public ICommand RemoveCommand { get; }

    public PresetItemViewModel(ShapePreset preset, SetupItemViewModel owner, ISetupHost host)
    {
        Preset        = preset;
        ShowCommand   = new RelayCommand(_ => host.ShowPreset(preset));
        RemoveCommand = new RelayCommand(_ => owner.RemovePreset(this));
    }
}

/// <summary>항목 카드가 부모(설정 화면 ViewModel)에게 요청하는 일</summary>
public interface ISetupHost
{
    bool IsHwpx { get; }
    /// <summary>이 문서와 다른 문서들에 있는 블록 값 이름 (처음 값·계산 고르기용)</summary>
    IReadOnlyList<string> KnownKeys();
    /// <summary>표의 줄 이름과 명단 직종을 서식에서 읽어 짝짓기 표를 채운다</summary>
    void LoadRoleMap(SetupItemViewModel item);
    void StartPick(SetupItemViewModel item, PickKind kind, string? columnKey = null);
    void RemoveItem(SetupItemViewModel item);
    void RevealItem(SetupItemViewModel item);
    void SavePreset(SetupItemViewModel item);
    void ShowPreset(ShapePreset preset);
}

/// <summary>체크해서 고르는 이름 하나 (계산에 쓸 항목, 짝짓기 직종)</summary>
public class KeyChoice : ViewModelBase
{
    private readonly Action _onToggled;
    public string Key { get; }
    /// <summary>지금은 없는 이름 (오타이거나 지운 항목) — 빨간 글씨</summary>
    public bool IsMissing { get; }
    public string Display => IsMissing ? $"{Key} (없음)" : Key;

    private bool _isChecked;
    public bool IsChecked { get => _isChecked; set { if (SetField(ref _isChecked, value)) _onToggled(); } }

    public KeyChoice(string key, bool isChecked, bool missing, Action onToggled)
    {
        Key = key; _isChecked = isChecked; IsMissing = missing; _onToggled = onToggled;
    }
}

/// <summary>짝짓기 표의 한 줄: 표의 줄 이름 ← 셀 명단 직종들</summary>
public class RoleMapRow
{
    public string RowName { get; }
    public ObservableCollection<KeyChoice> Jobs { get; }
    public RoleMapRow(string rowName, IEnumerable<KeyChoice> jobs) { RowName = rowName; Jobs = new(jobs); }
}

/// <summary>표 카드의 열 지정 버튼 하나</summary>
public record ColumnButton(string Key, string Text);

/// <summary>"적용할 시트" 선택 칩 하나</summary>
public class SheetChoice : ViewModelBase
{
    public string Name { get; }
    private bool _isChecked;
    public bool IsChecked { get => _isChecked; set => SetField(ref _isChecked, value); }
    public SheetChoice(string name, bool isChecked) { Name = name; _isChecked = isChecked; }
}
