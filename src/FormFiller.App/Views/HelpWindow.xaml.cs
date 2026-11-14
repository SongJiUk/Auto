using System.Windows;

namespace FormFiller.App.Views;

/// <summary>사용법 안내 창 (글만 있는 창이라 ViewModel 없음)</summary>
public partial class HelpWindow : Window
{
    public HelpWindow() => InitializeComponent();

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
