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
        Closed += (_, _) => CleanupExcel();
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
                    CleanupExcel();
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

    private void OpenExcel(string filePath, string sheetName)
    {
        CleanupExcel();

        var excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null)
        {
            MessageBox.Show("Excel이 설치되어 있지 않거나 등록되지 않았습니다.",
                "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        try
        {
            _xl = Activator.CreateInstance(excelType)!;
            _xl.Visible           = false;
            _xl.DisplayAlerts     = false;
            _xl.DisplayFormulaBar = false;
            _xl.DisplayStatusBar  = false;

            _wb = _xl.Workbooks.Open(filePath,
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
            CleanupExcel();
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

    private void CleanupExcel()
    {
        _pollTimer?.Stop();
        _pollTimer = null;

        ExcelHostControl.Content = null;
        _excelHost = null;

        try { _wb?.Close(false); }
        catch { }
        try { if (_wb != null) { Marshal.ReleaseComObject(_wb); _wb = null; } }
        catch { }

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
