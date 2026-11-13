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

    /// <summary>
    /// 기존 문서를 DocumentSetupWindow 로 열어 Excel 에서 클릭으로 다시 지정한다.
    /// 이 창에서 고치던 내용(저장 전)은 버려지므로, 필요하면 먼저 저장하라고 묻는다.
    /// </summary>
    private void EditDocument_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.Tag is not DocumentEditViewModel docVm) return;

        var doc = FormFiller.Core.Config.AppConfigLoader.Load().Documents.FirstOrDefault(d => d.Id == docVm.Id);
        if (doc == null)
        {
            MessageBox.Show("아직 저장되지 않은 문서입니다. 먼저 [저장]을 눌러 주세요.", "알림",
                            MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var setupWin = new DocumentSetupWindow(doc) { Owner = this };
        if (setupWin.ShowDialog() == true)
            DataContext = new SettingsViewModel();
    }
}
