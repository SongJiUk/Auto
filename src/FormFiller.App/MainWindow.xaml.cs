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
