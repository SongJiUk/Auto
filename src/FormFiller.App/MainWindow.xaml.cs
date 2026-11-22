using System.Windows;
using System.ComponentModel;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using FormFiller.App.ViewModels;
using FormFiller.App.Views;
using FormFiller.Core.Config;

namespace FormFiller.App;

/// <summary>
/// MainWindow 코드-비하인드.
/// MVVM 원칙에 따라 가능한 한 비워 두고,
/// WPF 이벤트 모델이 필요한 드래그&드롭 처리만 담는다.
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();

        // 단계를 바꾸면 오른쪽 화면을 맨 위부터 보여 준다
        if (DataContext is MainViewModel vm)
            vm.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(MainViewModel.CurrentStep)) BodyScroll.ScrollToTop();
            };

        // 공유 표시 띠 (다른 PC 에서 설정을 바꿨으면 이번에 켤 때 한 번 알린다)
        Loaded += (_, _) => RefreshSharedBar(showNotice: true);

        // 구글 드라이브 점검은 뒤에서 하고, 중요한 문제가 있을 때만 상단 띠로 알린다
        Loaded += async (_, _) => await CheckDriveAsync();

        // 처음 실행이면 사용법 창을 한 번 띄운다
        Loaded += (_, _) =>
        {
            var cfg = AppConfigLoader.Load();
            if (cfg.HelpShown) return;
            OpenHelp_Click(this, new RoutedEventArgs());
            cfg.HelpShown = true;
            AppConfigLoader.Save(cfg);
        };
    }

    // ── 표 칸 한 번 클릭으로 고치기 ─────────────────────────────────

    /// <summary>
    /// 고칠 수 있는 칸을 누르면 바로 입력 상태로 바꾸고 글자를 모두 골라 둔다.
    /// 그래서 "0" 을 지우지 않고 숫자를 바로 치면 된다.
    /// 삭제 버튼·누계처럼 고치는 모양이 없는 칸은 그대로 둔다.
    /// </summary>
    private void Cell_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGridCell cell || cell.IsEditing || cell.IsReadOnly) return;
        if (cell.Column is DataGridTemplateColumn { CellEditingTemplate: null }) return;

        // 널리 쓰는 WPF 방식: 칸에 포커스를 주고 줄을 미리 골라 두면,
        // DataGrid 가 "이미 고른 칸을 다시 누름"으로 보고 이 클릭에서 바로 입력 상태로 바꾼다.
        if (!cell.IsFocused) cell.Focus();
        if (FindParent<DataGridRow>(cell) is { IsSelected: false } row) row.IsSelected = true;

        // 입력 칸이 만들어진 뒤에 글자 전체 선택 (0 을 지우지 않고 바로 숫자를 치게)
        Dispatcher.BeginInvoke(() =>
        {
            if (cell.IsEditing && FindChild<TextBox>(cell) is { } box)
            {
                box.Focus();
                box.SelectAll();
            }
        }, DispatcherPriority.Input);
    }

    private static T? FindParent<T>(DependencyObject child) where T : DependencyObject
    {
        for (var p = VisualTreeHelper.GetParent(child); p != null; p = VisualTreeHelper.GetParent(p))
            if (p is T t) return t;
        return null;
    }

    private static T? FindChild<T>(DependencyObject parent) where T : DependencyObject
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var c = VisualTreeHelper.GetChild(parent, i);
            if (c is T t) return t;
            if (FindChild<T>(c) is { } found) return found;
        }
        return null;
    }

    /// <summary>입력을 바꾸고 파일을 만들지 않은 채 닫으면 한 번 묻는다.</summary>
    protected override void OnClosing(CancelEventArgs e)
    {
        if (DataContext is MainViewModel { HasUnsavedInput: true })
        {
            var r = MessageBox.Show(this,
                "입력한 내용으로 아직 파일을 만들지 않았습니다.\n닫으면 입력한 내용이 사라집니다. 닫을까요?",
                "닫기", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
            if (r != MessageBoxResult.Yes) e.Cancel = true;
        }
        base.OnClosing(e);
    }

    // ── 공유 표시 띠 ─────────────────────────────────────────────────

    /// <summary>
    /// "☁ 공유 설정 · 폴더 · 마지막 저장 PC·시각" 을 띠에 쓴다. 드라이브 밖이면 ⚠ 주황.
    /// showNotice 면 지난번에 켠 뒤 다른 PC 가 바꾼 경우 한 줄 알림을 덧붙이고, 본 시각을 이 PC 에 기록한다.
    /// </summary>
    private void RefreshSharedBar(bool showNotice)
    {
        try
        {
            var cfg      = AppConfigLoader.Load();
            string seen  = SharedConfig.DefaultSeenPath();
            var status   = SharedConfig.Describe(cfg, AppDomain.CurrentDomain.BaseDirectory,
                                                 Environment.MachineName, SharedConfig.ReadSeen(seen));

            SharedBarText.Text   = status.Text;
            SharedBarText.ToolTip = status.Text;
            SharedBar.Background = status.IsShared ? Brush("#EBF8FF") : Brush("#FEEBC8");
            SharedBarText.Foreground = status.IsShared ? Brush("#2C5282") : Brush("#7B341E");

            if (showNotice)
            {
                SharedNoticeText.Text       = status.Notice ?? "";
                SharedNoticeText.Visibility = status.Notice == null ? Visibility.Collapsed : Visibility.Visible;
                if (status.Notice != null) AppLog.Info(status.Notice);
            }
            if (cfg.LastSavedAt is { } at) SharedConfig.WriteSeen(seen, at);
        }
        catch (Exception ex)
        {
            AppLog.Error("공유 표시 띠 갱신 실패", ex);
        }
    }

    private static SolidColorBrush Brush(string hex) => (SolidColorBrush)new BrushConverter().ConvertFromString(hex)!;

    // ── 구글 드라이브 ───────────────────────────────────────────────

    private async Task CheckDriveAsync()
    {
        try
        {
            var cfg     = AppConfigLoader.Load();
            string bd   = AppDomain.CurrentDomain.BaseDirectory;
            var items   = await Task.Run(() => DriveChecker.Run(cfg, bd));
            // "프로그램 위치(드라이브 밖)"는 공유 표시 띠가 이미 ⚠ 로 보여 주므로 여기서는 뺀다 (같은 경고 두 줄 방지)
            var problem = items.FirstOrDefault(i => i.Important && i.State != CheckState.Ok && i.Title != "프로그램 위치");

            if (problem == null)
            {
                DriveBanner.Visibility = Visibility.Collapsed;
                return;
            }
            DriveBannerText.Text   = $"⚠ 구글 드라이브: {problem.Detail}  (눌러서 확인)";
            DriveBanner.Visibility = Visibility.Visible;
            AppLog.Info($"드라이브 점검 문제: {problem.Title} - {problem.Detail}");
        }
        catch (Exception ex)
        {
            AppLog.Error("시작 시 구글 드라이브 점검 실패", ex);
        }
    }

    private void DriveBanner_Click(object sender, MouseButtonEventArgs e) => OpenDrive_Click(sender, e);

    /// <summary>구글 드라이브 창 열기. 닫으면 폴더가 바뀌었을 수 있어 다시 읽고 다시 점검한다.</summary>
    private async void OpenDrive_Click(object sender, RoutedEventArgs e)
    {
        string outputBefore = AppConfigLoader.Load().OutputFolder;
        new DriveSettingsWindow { Owner = this }.ShowDialog();

        // 출력 폴더는 시작할 때 전일 값을 읽는 데 쓰므로, 바꿨으면 다시 켜야 새 폴더 기준으로 읽는다
        if (!string.Equals(outputBefore, AppConfigLoader.Load().OutputFolder, StringComparison.OrdinalIgnoreCase))
            MessageBox.Show(this,
                "출력 폴더를 바꿨습니다.\n전일 누계·어제 작업사항을 새 폴더에서 다시 읽으려면 프로그램을 다시 켜 주세요.",
                "구글 드라이브", MessageBoxButton.OK, MessageBoxImage.Information);

        // 입력해 둔 사진이 지워지지 않게 문서 목록은 다시 만들지 않고, 사진 폴더만 새 값으로 바꾼다
        if (DataContext is MainViewModel vm)
        {
            var cfg = AppConfigLoader.Load();
            foreach (var card in vm.DocInputCards)
                if (cfg.Documents.FirstOrDefault(d => d.Id == card.Config.Id) is { } doc)
                    card.UpdatePhotoFolder(doc.PhotoFolder);
        }
        await CheckDriveAsync();
    }

    /// <summary>사용법 창 열기</summary>
    private void OpenHelp_Click(object sender, RoutedEventArgs e)
        => new HelpWindow { Owner = this }.ShowDialog();

    /// <summary>설정 창 열기</summary>
    private void OpenSettings_Click(object sender, RoutedEventArgs e)
    {
        var win = new SettingsWindow { Owner = this };
        win.ShowDialog();

        // 설정에서 문서를 추가·편집했을 수 있으므로 문서 목록과 입력 칸을 다시 읽는다
        (DataContext as MainViewModel)?.ReloadDocuments();
        RefreshSharedBar(showNotice: false);   // 마지막 저장 시각 갱신 (알림은 켤 때만)
    }

    // ── 문서별 입력: 사진 칸 드래그&드롭 ──────────────────────────────

    private void DocPhoto_Drop(object sender, DragEventArgs e)
    {
        if (sender is FrameworkElement { DataContext: FieldInputViewModel item }
            && e.Data.GetData(DataFormats.FileDrop) is string[] { Length: > 0 } files)
        {
            item.SetPhoto(files[0]);
            e.Handled = true;
        }
    }

    /// <summary>사진 구역 카드의 사진 칸에 파일을 끌어다 놓았을 때</summary>
    private void BlockPhoto_Drop(object sender, DragEventArgs e)
    {
        if (sender is FrameworkElement { DataContext: PhotoPickViewModel slot }
            && e.Data.GetData(DataFormats.FileDrop) is string[] { Length: > 0 } files)
        {
            slot.SetDropped(files[0]);
            e.Handled = true;
        }
    }

    // ── 사진 슬롯 드래그&드롭 ────────────────────────────────────────

    /// <summary>이미지 파일을 슬롯 위로 드래그할 때 커서 변경</summary>
    private void PhotoSlot_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
            ? DragDropEffects.Copy
            : DragDropEffects.None;
        e.Handled = true;
    }

    /// <summary>
    /// 이미지 파일을 슬롯에 드롭했을 때:
    /// Border의 DataContext가 PhotoSlotViewModel이면 HandleDroppedFile 호출
    /// </summary>
    private void PhotoSlot_Drop(object sender, DragEventArgs e)
    {
        if (sender is not Border border) return;
        if (border.DataContext is not PhotoSlotViewModel slot) return;
        if (e.Data.GetData(DataFormats.FileDrop) is not string[] files || files.Length == 0) return;

        slot.HandleDroppedFile(files[0]);
        e.Handled = true;
    }
}
