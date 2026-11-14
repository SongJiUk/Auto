using System.Windows;
using FormFiller.App.ViewModels;

namespace FormFiller.App.Views;

/// <summary>[파일 생성] 확인·결과 창. 내용은 GenerateViewModel 이 모두 갖고 있다.</summary>
public partial class GenerateWindow : Window
{
    public GenerateWindow(GenerateViewModel vm)
    {
        InitializeComponent();
        DataContext = vm;

        // 만드는 중에 창을 닫으면 결과를 볼 수 없으니 끝날 때까지 막는다
        Closing += (_, e) => { if (vm.IsRunning) e.Cancel = true; };
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
