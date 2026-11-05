using System.Windows;
using FormFiller.App.ViewModels.Settings;

namespace FormFiller.App.Views;

/// <summary>
/// SettingsWindow 코드-비하인드.
/// 비즈니스 로직은 SettingsViewModel에 있고, 여기서는 "새 문서 추가" 버튼만 처리한다.
/// </summary>
public partial class SettingsWindow : Window
{
    public SettingsWindow()
    {
        InitializeComponent();
    }

    /// <summary>
    /// DocumentSetupWindow(새 문서 추가)를 모달로 열고,
    /// 닫히면 SettingsViewModel을 새로 로드해 목록을 갱신한다.
    /// </summary>
    private void AddDocument_Click(object sender, RoutedEventArgs e)
    {
        var setupWin = new DocumentSetupWindow { Owner = this };
        setupWin.ShowDialog();

        // DocumentSetupWindow에서 저장했으면 config.json이 바뀌었으므로
        // DataContext를 새 ViewModel로 교체해 목록을 갱신한다.
        DataContext = new SettingsViewModel();
    }
}
