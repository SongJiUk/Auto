using System.Windows.Input;

namespace FormFiller.App.ViewModels;

/// <summary>
/// 작업사항 목록의 항목 하나.
/// 텍스트 편집 + 삭제 버튼을 갖는다.
/// </summary>
public class WorkItemViewModel : ViewModelBase
{
    private string _text = "";

    /// <summary>항목 내용 (예: "- 굴삭기 작업")</summary>
    public string Text
    {
        get => _text;
        set => SetField(ref _text, value);
    }

    /// <summary>이 항목을 목록에서 제거하는 커맨드. 부모 ViewModel이 주입한다.</summary>
    public ICommand RemoveCommand { get; }

    /// <param name="text">초기 텍스트</param>
    /// <param name="onRemove">삭제 버튼 클릭 시 호출될 콜백 (부모가 컬렉션에서 제거)</param>
    public WorkItemViewModel(string text, Action<WorkItemViewModel> onRemove)
    {
        Text = text;
        RemoveCommand = new RelayCommand(_ => onRemove(this));
    }
}
