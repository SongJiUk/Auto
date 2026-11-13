using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using FormFiller.App.ViewModels.Settings;
using FormFiller.Core.Config;

namespace FormFiller.App.Views;

/// <summary>
/// DocumentSetupWindow 코드-비하인드.
/// Excel COM 을 dynamic + Type.GetTypeFromProgID 로 직접 사용한다.
/// NuGet Interop 패키지 불필요 → office.dll 의존성 없음.
/// 셀 감지는 200ms DispatcherTimer 폴링으로 처리한다.
/// </summary>
public partial class DocumentSetupWindow : Window
{
    private dynamic?        _xl;
    private dynamic?        _wb;
    private ExcelHwndHost?  _excelHost;
    private DispatcherTimer? _pollTimer;
    private string          _lastPolledAddr = "";

    public DocumentSetupWindow()
    {
        InitializeComponent();

        var vm = (DocumentSetupViewModel)DataContext;
        vm.PropertyChanged += OnVmChanged;
        vm.SaveSucceeded   += (_, _) => DialogResult = true;
        // 창이 없어지기 "전에" Excel 창의 소유자 지정을 풀어 둔다 (ExcelHwndHost.Detach 설명 참고)
        Closing += (_, _) => _excelHost?.Detach();
        Closed  += (_, _) => CleanupExcel();

        // 창이 뜨고 한가해지면 Excel 을 미리 켜 둔다 (Excel 시작에 1~2초 걸리므로 파일 선택 전에 끝내 둔다)
        Loaded += (_, _) => Dispatcher.BeginInvoke(
            () => { try { EnsureExcel(); } catch { /* 실패해도 파일 열 때 다시 시도 */ } },
            DispatcherPriority.ApplicationIdle);
    }

    private void OnVmChanged(object? sender, PropertyChangedEventArgs e)
    {
        var vm = (DocumentSetupViewModel)DataContext;
        switch (e.PropertyName)
        {
            case nameof(DocumentSetupViewModel.FilePath):
                if (!string.IsNullOrEmpty(vm.FilePath) && File.Exists(vm.FilePath))
                    OpenExcel(vm.FilePath, vm.SelectedSheet ?? "");
                else
                    CloseWorkbook();
                break;

            case nameof(DocumentSetupViewModel.SelectedSheet):
                SwitchSheet(vm.SelectedSheet ?? "");
                break;

            case nameof(DocumentSetupViewModel.IsPickerActive):
                UpdatePickerBorder();
                _lastPolledAddr = "";   // 피커 모드 전환 시 마지막 셀 초기화
                break;
        }
    }

    private void UpdatePickerBorder()
    {
        var vm = (DocumentSetupViewModel)DataContext;
        PreviewBorder.BorderBrush     = vm.IsPickerActive
            ? System.Windows.Media.Brushes.DodgerBlue
            : System.Windows.Media.Brushes.LightGray;
        PreviewBorder.BorderThickness = vm.IsPickerActive
            ? new Thickness(2) : new Thickness(1);
    }

    // ── Excel 열기 ─────────────────────────────────────────────────

    /// <summary>
    /// Excel 인스턴스가 없으면 만든다. 이미 있으면 그대로 재사용한다.
    /// (파일을 바꿀 때마다 Excel 을 새로 켜면 매번 1~2초씩 더 걸린다)
    /// </summary>
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

            // ReadOnly: 원본 서식 파일을 실수로 덮어쓰지 않기 위해 읽기 전용으로 연다.
            //           읽기 전용이어도 셀 선택·스크롤은 되므로 필드 지정에는 문제없다.
            _wb = _xl!.Workbooks.Open(filePath,
                UpdateLinks: false,
                ReadOnly: true,
                IgnoreReadOnlyRecommended: true);

            SwitchSheet(sheetName);

            _xl.Visible     = true;
            _xl.WindowState = -4143; // xlNormal

            // 창 초기화 대기 후 임베드
            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            timer.Tick += (_, _) => { timer.Stop(); EmbedExcelWindow(); StartPolling(); };
            timer.Start();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"엑셀을 열 수 없습니다:\n{ex.Message}",
                "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            CloseWorkbook();
        }
    }

    private void EmbedExcelWindow()
    {
        if (_xl == null) return;
        try
        {
            IntPtr hwnd = new IntPtr((int)_xl.Hwnd);
            if (hwnd == IntPtr.Zero) return;

            _excelHost = new ExcelHwndHost();
            ExcelHostControl.Content = _excelHost;
            _excelHost.AttachExcel(hwnd);
        }
        catch { }
    }

    /// <summary>
    /// 열려 있는 통합문서를 닫고 임베드를 해제한다. Excel 프로세스는 다음 파일을 위해 켜 둔다.
    /// (Excel 2013 이후는 통합문서마다 창이 따로 있어서, 파일이 바뀌면 임베드할 창도 바뀐다)
    /// </summary>
    private void CloseWorkbook()
    {
        _pollTimer?.Stop();
        _pollTimer = null;
        _lastPolledAddr = "";

        // [순서] ① 떼기(소유자 해제) → ② 숨기기 → ③ 닫기.
        _excelHost?.Detach();
        ExcelHostControl.Content = null;
        _excelHost = null;

        // 떼어낸 창이 바탕화면에 보이지 않도록 바로 숨긴다 (다음 파일도 임베드 전까지 숨김 상태로 열림)
        try { if (_xl != null) _xl.Visible = false; }
        catch { }

        try { _wb?.Close(false); }
        catch { }
        try { if (_wb != null) Marshal.ReleaseComObject(_wb); }
        catch { }
        _wb = null;
    }

    private void SwitchSheet(string sheetName)
    {
        if (_wb == null || string.IsNullOrEmpty(sheetName)) return;
        try
        {
            foreach (dynamic ws in _wb.Sheets)
                if ((string)ws.Name == sheetName) { ws.Activate(); return; }
        }
        catch { }
    }

    // ── 셀 폴링 (피커 모드 활성 시 현재 선택 셀 감지) ──────────────

    private void StartPolling()
    {
        _pollTimer?.Stop();
        _pollTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
        _pollTimer.Tick += (_, _) => PollActiveCell();
        _pollTimer.Start();
    }

    private void PollActiveCell()
    {
        var vm = (DocumentSetupViewModel)DataContext;
        if (!vm.IsPickerActive || _xl == null) return;

        try
        {
            dynamic? cell = _xl.ActiveCell;
            if (cell == null) return;

            int row = (int)cell.Row;
            int col = (int)cell.Column;
            string addr = ColLetter(col) + row;

            if (addr == _lastPolledAddr) return; // 같은 셀이면 무시
            _lastPolledAddr = addr;

            string value = cell.Value2 is null ? "" : cell.Value2.ToString();
            vm.OnCellClicked(addr, value);
        }
        catch { }
    }

    // ── Excel 정리 ────────────────────────────────────────────────

    /// <summary>창을 닫을 때 호출: 통합문서를 닫고 Excel 프로세스까지 종료한다.</summary>
    private void CleanupExcel()
    {
        CloseWorkbook();

        try { _xl?.Quit(); }
        catch { }
        try { if (_xl != null) { Marshal.ReleaseComObject(_xl); _xl = null; } }
        catch { }

        GC.Collect();
        GC.WaitForPendingFinalizers();
        _lastPolledAddr = "";
    }

    // ── 테이블 설정 버튼 ──────────────────────────────────────────

    private void TableSetup_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not FieldAssignRow field) return;
        bool isEquip = field.Type == FieldTypes.EquipmentTable;
        string msg = isEquip
            ? $"장비 테이블\n시작행:{field.StartRow} / 끝행:{field.EndRow}\n장비명:{field.NameCol} / 전일:{field.PrevDayCol} / 금일:{field.TodayCol}"
            : $"인원 테이블\n시작행:{field.StartRow} / 끝행:{field.EndRow}\n직종:{field.RoleCol} / 전일누계:{field.PrevCumulativeCol} / 금일:{field.TodayCol}";
        MessageBox.Show(msg + "\n\n※ 오른쪽 패널에서 직접 수정하세요.",
            "테이블 설정", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    // ── 드래그 앤 드롭 ────────────────────────────────────────────

    private void Window_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
            ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void Window_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] files)
            ((DocumentSetupViewModel)DataContext).HandleDroppedFile(files);
    }

    // ── 유틸리티 ─────────────────────────────────────────────────

    private static string ColLetter(int col)
    {
        string s = "";
        while (col > 0) { s = (char)('A' + (col - 1) % 26) + s; col = (col - 1) / 26; }
        return s;
    }
}
