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
}

/// <summary>Excel 선택 상태</summary>
/// <param name="IsShapes">그림·글상자가 선택됨 (아니면 칸)</param>
/// <param name="ActiveCell">칸: 현재 칸 주소(예: "D12") / 그림: 그림 왼쪽 위 칸</param>
/// <param name="Row">선택 범위 첫 행 (1부터)</param>
/// <param name="Col">선택 범위 첫 열 (1부터)</param>
/// <param name="RowCount">선택 범위 행 수</param>
/// <param name="ShapeNames">선택된 그림·글상자 이름들</param>
/// <param name="ShapeIndexes">선택된 그림·글상자의 시트 안 순서 (0부터)</param>
public record ExcelSelection(bool IsShapes, string ActiveCell, int Row, int Col, int RowCount,
                             List<string> ShapeNames, List<int> ShapeIndexes)
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
}

// ════════════════════════════════════════════════════════════════════
// 항목 종류 (추가할 때 고르는 목록)
// ════════════════════════════════════════════════════════════════════

public record ItemTypeOption(string Icon, string Label, string Type, string? FixedId = null)
{
    public string Display => $"{Icon}  {Label}";
    public override string ToString() => Display;

    public bool AsksSuffix => Type is FieldTypes.Input or FieldTypes.Photo;
    public bool AsksTime   => Type == FieldTypes.Photo;

    public static readonly IReadOnlyList<ItemTypeOption> All =
    [
        new("✏️", "매일 입력하는 값 (글자·숫자)",       FieldTypes.Input),
        new("📷", "사진 (+ 옆 칸 측정값)",              FieldTypes.Photo),
        new("📍", "그림·글상자 위치 (장비·측정기 표시)", FieldTypes.ShapePosition),
        new("📅", "날짜 — 2026.11.05",                  FieldTypes.DateSimple),
        new("📅", "날짜 — 2026년 11월 05일 목요일",     FieldTypes.DateKorean),
        new("🌤", "날씨",                               FieldTypes.Text, "weather"),
        new("📝", "금일 작업사항 (여러 줄)",            FieldTypes.Multiline, "today_work"),
        new("📝", "명일 작업사항 (여러 줄)",            FieldTypes.Multiline, "tomorrow_work"),
        new("🚜", "장비 투입현황 표",                   FieldTypes.EquipmentTable),
        new("👷", "인원 투입현황 표",                   FieldTypes.PersonnelTable),
        new("🚜", "투입 장비명 (자동)",                 FieldTypes.ActiveEquipmentNames),
        new("👷", "TBM 인원 (자동)",                    FieldTypes.TbmPersonnel),
        new("✅", "참석율 (항상 100)",                  FieldTypes.AttendanceRate),
    ];

    public static ItemTypeOption For(string type, string id) =>
        All.FirstOrDefault(o => o.Type == type && (o.FixedId == null || o.FixedId == id))
        ?? All.FirstOrDefault(o => o.Type == type)
        ?? new("•", type, type);
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
    public string Label { get => _label; set => SetField(ref _label, value); }

    // ── 지정 정보 ────────────────────────────────────────────────────
    private string? _cell, _valueCell, _suffix, _time, _placeholder;
    public string? Cell        { get => _cell;        set { if (SetField(ref _cell, value)) Changed(); } }
    public string? ValueCell   { get => _valueCell;   set { if (SetField(ref _valueCell, value)) Changed(); } }
    public string? Suffix      { get => _suffix;      set { if (SetField(ref _suffix, value)) Changed(); } }
    public string? Time        { get => _time;        set { if (SetField(ref _time, value)) Changed(); } }
    public string? Placeholder { get => _placeholder; set { if (SetField(ref _placeholder, value)) Changed(); } }

    public List<string> Cells { get; private set; } = [];
    public List<string> ShapeNames { get; private set; } = [];
    public List<int> ShapeIndexes { get; private set; } = [];
    public ObservableCollection<PresetItemViewModel> Presets { get; } = [];

    // 표
    public int StartRow { get; private set; }
    public int EndRow   { get; private set; }
    public Dictionary<string, string> Cols { get; } = [];   // "name","prev","today","role","prevCum" → 열 문자

    private string _newPresetName = "";
    public string NewPresetName { get => _newPresetName; set => SetField(ref _newPresetName, value); }

    // ── 화면 표시용 ──────────────────────────────────────────────────
    public bool IsHwpx        => _host.IsHwpx;
    public bool IsExcel       => !_host.IsHwpx;
    public bool IsPhoto       => Type == FieldTypes.Photo;
    public bool IsShape       => Type == FieldTypes.ShapePosition && !IsHwpx;
    public bool IsEquipTable  => Type == FieldTypes.EquipmentTable && !IsHwpx;
    public bool IsPersonTable => Type == FieldTypes.PersonnelTable && !IsHwpx;
    public bool HasSuffix     => Type is FieldTypes.Input or FieldTypes.Photo;

    /// <summary>카드에 보여 줄 지정 내용 요약 (예: "사진 D12 · 값 E12 · 10:00")</summary>
    public string Summary => IsHwpx ? "" : Type switch
    {
        FieldTypes.Photo => Join(Cell is null ? null : $"사진 {Cell}", ValueCell is null ? null : $"값 {ValueCell}",
                                 string.IsNullOrEmpty(Time) ? null : Time),
        FieldTypes.ShapePosition => ShapeNames.Count == 0 ? "" : $"{string.Join(", ", ShapeNames)} · 저장된 위치 {Presets.Count}개",
        FieldTypes.Multiline => Cells.Count == 0 ? "" : $"{Cells[0]} ~ {Cells[^1]} ({Cells.Count}줄)",
        FieldTypes.EquipmentTable or FieldTypes.PersonnelTable =>
            StartRow == 0 ? "" : $"{StartRow}~{EndRow}행 · " + string.Join(" ", ColKeys.Select(k => $"{ColTitle(k)} {ColOf(k) ?? "?"}")),
        _ => Cell ?? "",
    };

    /// <summary>엑셀에서 지정이 끝났는지 (카드의 ✓ / "지정 필요")</summary>
    public bool IsComplete => IsHwpx ? !string.IsNullOrWhiteSpace(Placeholder) : Type switch
    {
        FieldTypes.Photo         => Cell != null,
        FieldTypes.ShapePosition => ShapeNames.Count > 0,
        FieldTypes.Multiline     => Cells.Count > 0,
        FieldTypes.EquipmentTable or FieldTypes.PersonnelTable => StartRow > 0 && ColKeys.All(k => ColOf(k) != null),
        _ => Cell != null,
    };

    /// <summary>표 종류별로 지정할 열 (순서대로 안내)</summary>
    public string[] ColKeys => Type == FieldTypes.EquipmentTable ? ["name", "prev", "today"]
                             : Type == FieldTypes.PersonnelTable ? ["role", "prevCum", "today"] : [];

    public static string ColTitle(string key) => key switch
    {
        "name" => "장비명", "prev" => "전일", "today" => "금일", "role" => "직종", "prevCum" => "전일누계", _ => key,
    };

    public string? ColOf(string key) => Cols.TryGetValue(key, out var v) ? v : null;

    /// <summary>표 카드의 열 버튼들 (예: "장비명 B", "전일 ?")</summary>
    public List<ColumnButton> ColumnButtons =>
        ColKeys.Select(k => new ColumnButton(k, $"{ColTitle(k)} {ColOf(k) ?? "?"}")).ToList();

    // ── 커맨드 ───────────────────────────────────────────────────────
    public ICommand PickCommand       { get; }
    public ICommand PickValueCommand  { get; }
    public ICommand PickColumnCommand { get; }
    public ICommand RemoveCommand     { get; }
    public ICommand RevealCommand     { get; }
    public ICommand SavePresetCommand { get; }

    public SetupItemViewModel(ISetupHost host, string id, string label, string type)
    {
        _host  = host;
        Id     = id;
        _label = label;
        Type   = type;
        Option = ItemTypeOption.For(type, id);

        PickCommand       = new RelayCommand(_ => _host.StartPick(this, MainPickKind));
        PickValueCommand  = new RelayCommand(_ => _host.StartPick(this, PickKind.ValueCell));
        PickColumnCommand = new RelayCommand(p => _host.StartPick(this, PickKind.Column, p as string));
        RemoveCommand     = new RelayCommand(_ => _host.RemoveItem(this));
        RevealCommand     = new RelayCommand(_ => _host.RevealItem(this));
        SavePresetCommand = new RelayCommand(_ => _host.SavePreset(this));
    }

    /// <summary>이 항목을 처음 지정할 때 Excel 에서 무엇을 골라야 하는지</summary>
    public PickKind MainPickKind => Type switch
    {
        FieldTypes.Photo         => PickKind.Picture,
        FieldTypes.ShapePosition => PickKind.Shapes,
        FieldTypes.Multiline     => PickKind.Lines,
        FieldTypes.EquipmentTable or FieldTypes.PersonnelTable => PickKind.Rows,
        _ => PickKind.Cell,
    };

    // ── 값 설정 (ViewModel 에서 호출) ────────────────────────────────

    public void SetShapes(List<string> names, List<int> indexes) { ShapeNames = names; ShapeIndexes = indexes; Changed(); }
    public void SetLines(List<string> cells)  { Cells = cells; Changed(); }
    public void SetRows(int start, int end)   { StartRow = start; EndRow = end; Changed(); }
    public void SetColumn(string key, string col) { Cols[key] = col; Changed(); }

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
    }

    /// <summary>엑셀/한글 전환 시 표시 갱신</summary>
    public void RefreshMode()
    {
        foreach (var n in new[] { nameof(IsHwpx), nameof(IsExcel), nameof(IsShape), nameof(IsEquipTable),
                                  nameof(IsPersonTable), nameof(Summary), nameof(IsComplete) })
            OnPropertyChanged(n);
    }

    private static string Join(params string?[] parts) => string.Join(" · ", parts.Where(p => !string.IsNullOrEmpty(p)));

    // ── config 변환 ──────────────────────────────────────────────────

    public static SetupItemViewModel FromConfig(ISetupHost host, FieldConfig f)
    {
        var item = new SetupItemViewModel(host, f.Id, f.Label, f.Type)
        {
            _cell = f.Cell, _valueCell = f.ValueCell, _suffix = f.Suffix, _time = f.Time,
            _placeholder = f.Placeholder,
        };
        item.Cells      = f.Cells ?? [];
        item.ShapeNames = f.ShapeNames ?? [];
        item.ShapeIndexes = f.ShapeIndexes ?? [];
        item.StartRow   = f.StartRow;
        item.EndRow     = f.EndRow;
        void Col(string k, string? v) { if (!string.IsNullOrEmpty(v)) item.Cols[k] = v; }
        Col("name", f.NameCol); Col("prev", f.PrevDayCol); Col("today", f.TodayCol);
        item._specCol = f.SpecCol;   // 화면에서 고르지 않는 값 — 그대로 보존
        Col("role", f.RoleCol); Col("prevCum", f.PrevCumulativeCol);
        foreach (var p in f.Presets ?? []) item.Presets.Add(new PresetItemViewModel(p, item, host));
        return item;
    }

    private string? _specCol;

    public FieldConfig ToConfig()
    {
        var f = new FieldConfig { Id = Id, Label = Label.Trim(), Type = Type };
        if (IsHwpx)
        {
            f.Placeholder = string.IsNullOrWhiteSpace(Placeholder) ? null : Placeholder.Trim();
            f.Suffix = NullIfEmpty(Suffix);
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
    void StartPick(SetupItemViewModel item, PickKind kind, string? columnKey = null);
    void RemoveItem(SetupItemViewModel item);
    void RevealItem(SetupItemViewModel item);
    void SavePreset(SetupItemViewModel item);
    void ShowPreset(ShapePreset preset);
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
