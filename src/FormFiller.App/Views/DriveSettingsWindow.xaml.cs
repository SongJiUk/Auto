using System.Windows;
using FormFiller.App.ViewModels;

namespace FormFiller.App.Views;

/// <summary>구글 드라이브 점검·폴더 설정 창. 내용은 DriveSettingsViewModel 이 맡는다.</summary>
public partial class DriveSettingsWindow : Window
{
    public DriveSettingsWindow()
    {
        InitializeComponent();

        // 사진 칸이 있는 문서가 없으면 안내 한 줄만 보인다
        if (DataContext is DriveSettingsViewModel vm)
        {
            void Update() => NoPhotoText.Visibility = vm.PhotoFolders.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            vm.PhotoFolders.CollectionChanged += (_, _) => Update();
            Update();
        }
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
