using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using FormFiller.Core.Config;

namespace FormFiller.App.ViewModels.Settings;

/// <summary>
/// 설정 창(SettingsWindow)의 ViewModel.
///
/// config.json에 있는 모든 문서를 DocumentEditViewModel로 감싸
/// Documents 컬렉션으로 노출한다.
/// 저장 시 각 DocumentEditViewModel.ToConfig()로 변환 후 config.json에 덮어쓴다.
/// </summary>
public class SettingsViewModel : ViewModelBase
{
    /// <summary>편집 중인 문서 목록 (ItemsControl에 바인딩)</summary>
    public ObservableCollection<DocumentEditViewModel> Documents { get; } = [];

    /// <summary>모든 변경을 config.json에 저장한다</summary>
    public ICommand SaveCommand { get; }

    public SettingsViewModel()
    {
        var config = AppConfigLoader.Load();

        foreach (var doc in config.Documents)
            Documents.Add(new DocumentEditViewModel(doc, RemoveDocument));

        SaveCommand = new RelayCommand(_ => OnSave());
    }

    // ── 메서드 ──────────────────────────────────────────────────────

    /// <summary>문서 삭제 버튼에서 호출된다. 확인 후 목록에서 제거한다.</summary>
    private void RemoveDocument(DocumentEditViewModel doc)
    {
        var result = MessageBox.Show(
            $"'{doc.Name}' 문서를 삭제할까요?",
            "문서 삭제", MessageBoxButton.YesNo, MessageBoxImage.Question);

        if (result == MessageBoxResult.Yes)
            Documents.Remove(doc);
    }

    /// <summary>
    /// 현재 Documents 목록 전체를 config.json에 저장한다.
    /// AppConfigLoader.Load()로 outputFolder 등 다른 값은 보존한다.
    /// </summary>
    private void OnSave()
    {
        try
        {
            var config = AppConfigLoader.Load();
            config.Documents = Documents.Select(d => d.ToConfig()).ToList();
            AppConfigLoader.Save(config);

            MessageBox.Show("저장되었습니다.", "저장 완료",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"저장 오류\n{ex.Message}", "오류",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
