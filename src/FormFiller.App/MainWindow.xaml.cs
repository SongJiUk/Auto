using System.Windows;
using System.Windows.Controls;
using FormFiller.App.ViewModels;
using FormFiller.App.Views;

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
    }

    /// <summary>설정 창 열기</summary>
    private void OpenSettings_Click(object sender, RoutedEventArgs e)
    {
        var win = new SettingsWindow { Owner = this };
        win.ShowDialog();
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
