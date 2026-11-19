using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using FormFiller.App.ViewModels.Settings;
using FormFiller.Core.Config;

namespace FormFiller.App.Views;

/// <summary>
/// DocumentSetupWindow 코드-비하인드.
/// - Excel COM 을 dynamic + Type.GetTypeFromProgID 로 직접 사용한다 (Interop 패키지 불필요).
/// - IExcelBridge 를 구현해 ViewModel 이 "지금 선택된 칸/그림", "그림 위치" 등을 물어볼 수 있게 한다.
/// - 200ms 타이머로 Excel 선택이 바뀌었는지 살핀다 (Excel 이벤트를 COM 으로 받는 것보다 단순하고 안정적).
/// </summary>
public partial class DocumentSetupWindow : Window, IExcelBridge
{
    private dynamic?         _xl;
    private dynamic?         _wb;
    private ExcelHwndHost?   _excelHost;
    private DispatcherTimer? _pollTimer;
    private string?          _pickBaseline;   // 고르기 시작할 때의 선택 (이후 "바뀐" 선택만 지정)
    private string?          _freeBaseline;   // 고르는 중이 아닐 때 마지막으로 본 선택

    private DocumentSetupViewModel Vm => (DocumentSetupViewModel)DataContext;

    /// <summary>새 문서 추가</summary>
    public DocumentSetupWindow() : this(null) { }

    /// <summary>existing 이 있으면 그 문서를 편집한다.</summary>
    public DocumentSetupWindow(DocumentConfig? existing)
    {
        InitializeComponent();

        Vm.Bridge = this;
        Vm.PropertyChanged += OnVmChanged;
        Vm.SaveSucceeded   += (_, _) => DialogResult = true;
        Vm.PickStarted     += (_, _) => _pickBaseline = SafeSelection()?.Signature;
        // 칸을 골라 새 항목을 만들면 그 카드가 보이게 스크롤
        Vm.ItemCreated     += (_, item) => Dispatcher.BeginInvoke(
            () => (ItemsList.ItemContainerGenerator.ContainerFromItem(item) as FrameworkElement)?.BringIntoView(),
            DispatcherPriority.Loaded);

        // 창이 없어지기 "전에" Excel 창의 소유자 지정을 풀어 둔다 (ExcelHwndHost.Detach 설명 참고)
        Closing += (_, _) => _excelHost?.Detach();
        Closed  += (_, _) => CleanupExcel();

        Loaded += (_, _) =>
        {
            if (existing != null) Vm.LoadExisting(existing);
            // 한가해지면 Excel 을 미리 켜 둔다 (Excel 시작에 1~2초 걸림)
            Dispatcher.BeginInvoke(() => { try { if (Vm.IsExcel) EnsureExcel(); } catch { } },
                                   DispatcherPriority.ApplicationIdle);
        };
    }

    private void OnVmChanged(object? sender, PropertyChangedEventArgs e)
    {
        switch (e.PropertyName)
        {
            case nameof(DocumentSetupViewModel.FilePath):
                if (Vm.IsExcel && File.Exists(Vm.FilePath)) OpenExcel(Vm.FilePath, Vm.SelectedSheet);
                else CloseWorkbook();
                break;

            case nameof(DocumentSetupViewModel.SelectedSheet):
                SwitchSheet(Vm.SelectedSheet);
                break;

            case nameof(DocumentSetupViewModel.IsPicking):
                PreviewBorder.BorderBrush     = Vm.IsPicking ? System.Windows.Media.Brushes.DodgerBlue
                                                             : new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0xCB, 0xD5, 0xE0));
                PreviewBorder.BorderThickness = new Thickness(Vm.IsPicking ? 3 : 1);
                // 고르기가 끝나면 방금 고른 칸으로 "이 칸은?" 이 다시 뜨지 않게 지금 선택을 기억해 둔다
                if (!Vm.IsPicking) _freeBaseline = SafeSelection()?.Signature;
                break;
        }
    }

    // ── Excel 열기 / 닫기 ────────────────────────────────────────────

    /// <summary>Excel 인스턴스가 없으면 만든다 (파일을 바꿔도 재사용 → 빠름).</summary>
    private void EnsureExcel()
    {
        if (_xl != null) return;
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel이 설치되어 있지 않거나 등록되지 않았습니다.");

        _xl = Activator.CreateInstance(excelType)!;
        _xl.Visible           = false;
        _xl.DisplayAlerts     = false;
        _xl.DisplayFormulaBar = false;
        _xl.DisplayStatusBar  = false;
    }

    private void OpenExcel(string filePath, string sheetName)
    {
        CloseWorkbook();
        try
        {
            EnsureExcel();

            // ReadOnly: 원본 서식을 실수로 덮어쓰지 않도록. 칸 선택·그림 옮기기(미리보기)는 그대로 된다.
            _wb = _xl!.Workbooks.Open(filePath, UpdateLinks: false, ReadOnly: true, IgnoreReadOnlyRecommended: true);
            SwitchSheet(sheetName);

            _xl.Visible     = true;
            _xl.WindowState = -4143; // xlNormal

            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            timer.Tick += (_, _) => { timer.Stop(); EmbedExcelWindow(); StartPolling(); Vm.RefreshHighlights(); };
            timer.Start();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"엑셀을 열 수 없습니다:\n{ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            CloseWorkbook();
        }
    }

    private void EmbedExcelWindow()
    {
        if (_xl == null || _wb == null) return;
        try
        {
            IntPtr hwnd = new((int)_xl.Hwnd);
            if (hwnd == IntPtr.Zero) return;
            _excelHost = new ExcelHwndHost();
            ExcelHostControl.Content = _excelHost;
            _excelHost.AttachExcel(hwnd);
        }
        catch { }
    }

    /// <summary>통합문서를 닫는다. Excel 프로세스는 다음 파일을 위해 켜 둔다.</summary>
    private void CloseWorkbook()
    {
        _pollTimer?.Stop();
        _pollTimer = null;
        Vm.CancelPick();

        // [순서] ① 떼기(소유자 해제) → ② 숨기기 → ③ 닫기(저장 안 함)
        _painted.Clear();   // 통합문서를 닫으면(저장 안 함) 칠한 색도 같이 사라진다
        _excelHost?.Detach();
        ExcelHostControl.Content = null;
        _excelHost = null;

        try { if (_xl != null) _xl.Visible = false; } catch { }
        try { _wb?.Close(false); } catch { }
        try { if (_wb != null) Marshal.ReleaseComObject(_wb); } catch { }
        _wb = null;
    }

    private void SwitchSheet(string sheetName)
    {
        if (_wb == null || string.IsNullOrEmpty(sheetName)) return;
        try
        {
            if ((string)_xl!.ActiveSheet.Name == sheetName) return;
            foreach (dynamic ws in _wb.Sheets)
                if ((string)ws.Name == sheetName) { ws.Activate(); return; }
        }
        catch { }
    }

    private void CleanupExcel()
    {
        CloseWorkbook();
        try { _xl?.Quit(); } catch { }
        try { if (_xl != null) { Marshal.ReleaseComObject(_xl); _xl = null; } } catch { }
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    // ── 선택 살피기 (200ms) ──────────────────────────────────────────

    private void StartPolling()
    {
        _pollTimer?.Stop();
        _pollTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
        _pollTimer.Tick += (_, _) => Poll();
        _pollTimer.Start();
    }

    private void Poll()
    {
        if (_wb == null) return;

        // Excel 아래쪽 시트 탭으로 시트를 바꾸면 화면 상태도 맞춘다
        try
        {
            string active = (string)_xl!.ActiveSheet.Name;
            if (active != Vm.SelectedSheet) Vm.SelectedSheet = active;
        }
        catch { return; }   // Excel 이 바쁨(셀 편집 중 등) → 다음에

        if (!Vm.IsPicking)
        {
            // 고르는 중이 아니면: 칸을 고를 때마다 "이 칸은?" 선택지를 보여 준다
            var free = SafeSelection();
            if (free == null || free.Signature == _freeBaseline) return;
            _freeBaseline = free.Signature;
            Vm.OnFreeSelection(free);
            return;
        }
        _freeBaseline = null;

        var sel = SafeSelection();
        if (sel == null || sel.Signature == _pickBaseline) return;
        _pickBaseline = sel.Signature;
        Vm.OnExcelSelectionChanged(sel);
    }

    private ExcelSelection? SafeSelection()
    {
        try { return GetSelection(); } catch { return null; }
    }

    // ── IExcelBridge 구현 ────────────────────────────────────────────

    public ExcelSelection? GetSelection()
    {
        if (_wb == null || _xl == null) return null;
        dynamic sel = _xl.Selection;
        if (sel == null) return null;

        // 그림·글상자가 선택되어 있으면 ShapeRange 가 있다 (칸 선택이면 없어서 예외)
        try
        {
            dynamic sr = sel.ShapeRange;
            int n = (int)sr.Count;
            var names = new List<string>();
            var order = new List<int>();
            for (int i = 1; i <= n; i++)
            {
                names.Add((string)sr.Item(i).Name);
                order.Add((int)sr.Item(i).ZOrderPosition - 1);   // 시트 안 그림 순서 (파일의 그림 순서와 같다)
            }
            dynamic first = sr.Item(1);
            dynamic tl = first.TopLeftCell, br = first.BottomRightCell;
            // 그림이 덮는 칸 범위 (사진 구역 지정용). 그림 끝이 칸 경계에 딱 맞으면 BottomRightCell 이 다음 칸이라 한 칸 뺀다.
            int lastRow = (int)br.Row, lastCol = (int)br.Column;
            if ((double)first.Top + (double)first.Height - (double)br.Top < 2 && lastRow > (int)tl.Row) lastRow--;
            if ((double)first.Left + (double)first.Width - (double)br.Left < 2 && lastCol > (int)tl.Column) lastCol--;
            return new ExcelSelection(true, (string)tl.Address(false, false), (int)tl.Row, (int)tl.Column,
                                      lastRow - (int)tl.Row + 1, names, order, lastCol - (int)tl.Column + 1);
        }
        catch { /* 칸 선택 */ }

        // 병합된 칸 하나를 고른 것도 "칸 하나"로 본다 (선택 범위 = 그 칸의 병합 범위)
        int rows = (int)sel.Rows.Count, cols = (int)sel.Columns.Count;
        bool one = rows * cols == 1;
        if (!one) try { one = (string)sel.Address(false, false) == (string)_xl.ActiveCell.MergeArea.Address(false, false); } catch { }
        return new ExcelSelection(false, (string)_xl.ActiveCell.Address(false, false),
                                  (int)sel.Row, (int)sel.Column, rows, [], [], cols, one);
    }

    // ── 정한 칸 색 표시 ──────────────────────────────────────────────
    // Excel 의 색 = BGR 숫자. 연한 초록 RGB(198,239,206)
    private const int HighlightBgr = 206 * 65536 + 239 * 256 + 198;
    private const int XlNone = -4142;

    /// <summary>칠하기 전 원래 색 (시트, 범위, 색(null=섞임/없음), 무늬) — 칠한 순서대로</summary>
    private readonly List<(string Sheet, string Range, object? Color, int Pattern)> _painted = [];

    /// <summary>
    /// 앞에 칠한 칸을 원래 색으로 돌리고(거꾸로), 지금 시트에 새로 칠한다.
    /// 통합문서는 읽기 전용으로 열었고 닫을 때 저장하지 않으므로 서식 파일은 바뀌지 않는다.
    /// </summary>
    public void Highlight(IEnumerable<string> ranges)
    {
        if (_wb == null || _xl == null) return;
        dynamic wb = _wb!, xl = _xl!;
        bool screen = true;
        try { screen = xl.ScreenUpdating; xl.ScreenUpdating = false; } catch { }
        try
        {
            for (int i = _painted.Count - 1; i >= 0; i--)
            {
                var (sheet, range, color, pattern) = _painted[i];
                try
                {
                    dynamic interior = wb.Sheets[sheet].Range(range).Interior;
                    if (color == null || pattern == XlNone) interior.ColorIndex = XlNone;
                    else { interior.Pattern = pattern; interior.Color = color; }
                }
                catch { /* 시트가 없어졌거나 잘못된 주소 */ }
            }
            _painted.Clear();

            dynamic ws = xl.ActiveSheet;
            string name = ws.Name;
            foreach (string r in ranges)
            {
                try
                {
                    dynamic interior = ws.Range(r).Interior;
                    object? color = interior.Color is double d ? d : null;   // 섞인 색이면 DBNull
                    int pattern = interior.Pattern is int p ? p : XlNone;
                    _painted.Add((name, r, color, pattern));
                    interior.Color = HighlightBgr;
                }
                catch { /* 잘못된 주소는 건너뜀 */ }
            }
        }
        finally { try { xl.ScreenUpdating = screen; } catch { } }
    }

    /// <summary>지금 시트 칸에 보이는 글자 (병합 칸은 왼쪽 위 칸에 글자가 있다)</summary>
    public string GetCellText(int row, int col)
    {
        try
        {
            dynamic ws = _xl!.ActiveSheet;
            return (string)(ws.Cells[row, col].Text ?? "");
        }
        catch { return ""; }
    }

    /// <summary>시트 안 순서(0부터)로 그림·글상자를 찾는다.</summary>
    private static dynamic? ShapeAt(dynamic ws, int index)
    {
        foreach (dynamic sh in ws.Shapes)
            if ((int)sh.ZOrderPosition == index + 1) return sh;
        return null;
    }

    /// <summary>1pt = 12,700 EMU (엑셀 내부 위치 단위)</summary>
    private const double EmuPerPoint = 12700.0;

    public List<ShapePlacement> ReadPlacements(IEnumerable<int> shapeIndexes)
    {
        var list = new List<ShapePlacement>();
        if (_wb == null) return list;
        dynamic ws = _xl!.ActiveSheet;

        foreach (int index in shapeIndexes)
        {
            try
            {
                dynamic? sh = ShapeAt(ws, index);
                if (sh == null) continue;
                string name = sh.Name;
                double left = sh.Left, top = sh.Top, right = left + (double)sh.Width, bottom = top + (double)sh.Height;
                dynamic tl = sh.TopLeftCell, br = sh.BottomRightCell;
                list.Add(new ShapePlacement
                {
                    Name       = name,
                    Index      = index,
                    FromCol    = (int)tl.Column - 1, FromColOff = (long)((left - (double)tl.Left) * EmuPerPoint),
                    FromRow    = (int)tl.Row - 1,    FromRowOff = (long)((top  - (double)tl.Top)  * EmuPerPoint),
                    ToCol      = (int)br.Column - 1, ToColOff   = (long)((right  - (double)br.Left) * EmuPerPoint),
                    ToRow      = (int)br.Row - 1,    ToRowOff   = (long)((bottom - (double)br.Top)  * EmuPerPoint),
                });
            }
            catch { /* 이 시트에 그 이름의 그림이 없음 */ }
        }
        return list;
    }

    public void ApplyPlacements(IEnumerable<ShapePlacement> placements)
    {
        if (_wb == null) return;
        try
        {
            dynamic ws = _xl!.ActiveSheet;
            foreach (var p in placements)
            {
                try
                {
                    dynamic? sh = p.Index >= 0 ? ShapeAt(ws, p.Index) : ws.Shapes.Item(p.Name);
                    if (sh == null) continue;
                    dynamic c  = ws.Cells(p.FromRow + 1, p.FromCol + 1);
                    sh.Left = (double)c.Left + p.FromColOff / EmuPerPoint;
                    sh.Top  = (double)c.Top  + p.FromRowOff / EmuPerPoint;
                }
                catch { }
            }
        }
        catch { }
    }

    public void Reveal(string? cell, IEnumerable<int>? shapeIndexes)
    {
        if (_wb == null) return;
        try
        {
            dynamic ws = _xl!.ActiveSheet;
            var names = new List<object>();
            foreach (int i in shapeIndexes ?? [])
                if (ShapeAt(ws, i) is { } sh) names.Add((string)sh.Name);

            if (names.Count > 0) ws.Shapes.Range(names.ToArray()).Select();
            else if (!string.IsNullOrEmpty(cell)) ws.Range(cell).Select();
        }
        catch { }
    }

    /// <summary>시험·자동화용: 이름으로 그림을 선택한다 (사용자 클릭과 같은 효과).</summary>
    public void SelectShapesByName(params string[] names)
    {
        try { _xl!.ActiveSheet.Shapes.Range(names.Cast<object>().ToArray()).Select(); } catch { }
    }

    // ── 드래그 앤 드롭 ────────────────────────────────────────────

    private void Window_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void Window_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] files)
            Vm.HandleDroppedFile(files);
    }
}
