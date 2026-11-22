using System.Windows;
using FormFiller.App.ViewModels.Settings;
using FormFiller.Core.Config;

namespace FormFiller.App.Views;

/// <summary>
/// 문서 관리 창 코드-비하인드. 로직은 SettingsViewModel 에 있고,
/// 여기서는 [+ 새 문서]·[✏ 칸 편집] (다른 창을 여는 버튼)만 처리한다.
/// </summary>
public partial class SettingsWindow : Window
{
    public SettingsWindow()
    {
        InitializeComponent();
    }

    private SettingsViewModel Vm => (SettingsViewModel)DataContext;

    /// <summary>새 문서 등록 창을 열고, 닫히면 목록을 다시 읽는다 (새로 만든 문서를 고른 상태로).</summary>
    private void AddDocument_Click(object sender, RoutedEventArgs e)
    {
        if (!Vm.Save(showDone: false)) return;   // 이 창에서 고친 것을 잃지 않게 먼저 저장
        var before = AppConfigLoader.Load().Documents.Select(d => d.Id).ToHashSet();

        new DocumentSetupWindow { Owner = this }.ShowDialog();

        string? added = AppConfigLoader.Load().Documents.Select(d => d.Id).FirstOrDefault(id => !before.Contains(id));
        DataContext = new SettingsViewModel(added);
    }

    /// <summary>고른 문서를 칸 편집 창(Excel 클릭)으로 연다. 이 창에서 고친 것은 먼저 저장한다.</summary>
    private void EditDocument_Click(object sender, RoutedEventArgs e)
    {
        if (Vm.SelectedDocument is not { } docVm) return;
        if (!Vm.Save(showDone: false)) return;

        var doc = AppConfigLoader.Load().Documents.FirstOrDefault(d => d.Id == docVm.Id);
        if (doc == null) return;

        new DocumentSetupWindow(doc) { Owner = this }.ShowDialog();
        DataContext = new SettingsViewModel(docVm.Id);
    }
}
