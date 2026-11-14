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
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
