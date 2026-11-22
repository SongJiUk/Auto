using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using FormFiller.Core.Config;

namespace FormFiller.App.ViewModels.Settings;

/// <summary>
/// 문서 관리 창(SettingsWindow)의 ViewModel.
/// 왼쪽 = 문서 목록(Documents), 오른쪽 = 고른 문서(SelectedDocument)의 기본 정보 + 항목 요약.
/// 저장 시 각 DocumentEditViewModel.ToConfig()로 변환 후 config.json에 덮어쓴다.
/// </summary>
public class SettingsViewModel : ViewModelBase
{
    public ObservableCollection<DocumentEditViewModel> Documents { get; } = [];

    private DocumentEditViewModel? _selectedDocument;
    public DocumentEditViewModel? SelectedDocument
    {
        get => _selectedDocument;
        set { if (SetField(ref _selectedDocument, value)) OnPropertyChanged(nameof(HasSelection)); }
    }

    public bool HasSelection => _selectedDocument != null;

    public ICommand SaveCommand   { get; }
    public ICommand DeleteCommand { get; }

    /// <summary>XAML 에서 만드는 기본 생성자 (첫 문서를 고른다)</summary>
    public SettingsViewModel() : this(null) { }

    /// <param name="selectId">처음에 고를 문서 Id (없으면 첫 문서)</param>
    public SettingsViewModel(string? selectId)
    {
        var config = AppConfigLoader.Load();
        foreach (var doc in config.Documents)
            Documents.Add(new DocumentEditViewModel(doc));

        SelectedDocument = Documents.FirstOrDefault(d => d.Id == selectId) ?? Documents.FirstOrDefault();

        SaveCommand   = new RelayCommand(_ => OnSave());
        DeleteCommand = new RelayCommand(_ => RemoveSelected());
    }

    private void RemoveSelected()
    {
        if (_selectedDocument is not { } doc) return;
        var result = MessageBox.Show(
            $"'{doc.Name}' 문서를 목록에서 지울까요?\n(서식 파일·만든 파일은 지워지지 않습니다. [저장]을 눌러야 반영됩니다.)",
            "문서 삭제", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes) return;

        int index = Documents.IndexOf(doc);
        Documents.Remove(doc);
        SelectedDocument = Documents.Count == 0 ? null : Documents[Math.Min(index, Documents.Count - 1)];
    }

    /// <summary>현재 Documents 목록 전체를 config.json에 저장한다 (outputFolder 등 다른 값은 보존).</summary>
    public bool Save(bool showDone = true)
    {
        try
        {
            var config = AppConfigLoader.Load();
            config.Documents = Documents.Select(d => d.ToConfig()).ToList();
            AppConfigLoader.Save(config);
            if (showDone)
                MessageBox.Show("저장되었습니다.", "저장 완료", MessageBoxButton.OK, MessageBoxImage.Information);
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"저장 오류\n{ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
    }

    private void OnSave() => Save();
}
