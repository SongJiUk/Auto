namespace FormFiller.App.ViewModels;

/// <summary>
/// 메인 화면 왼쪽 단계 목록의 한 줄 (예: "② 작업사항").
/// 오늘 만들 문서에 필요 없는 단계는 IsVisible=false 로 숨긴다.
/// IsDone 이면 번호 대신 ✔ 를 보여 준다.
/// </summary>
public class StepItem : ViewModelBase
{
    public string Key   { get; }
    public string Title { get; }

    public StepItem(string key, string title)
    {
        Key   = key;
        Title = title;
    }

    private int _number;
    /// <summary>보이는 단계 중 몇 번째인지 (숨긴 단계는 건너뛰고 1, 2, 3 …)</summary>
    public int Number { get => _number; set { if (SetField(ref _number, value)) OnPropertyChanged(nameof(Badge)); } }

    private bool _isVisible = true;
    public bool IsVisible { get => _isVisible; set => SetField(ref _isVisible, value); }

    private bool _isDone;
    public bool IsDone { get => _isDone; set { if (SetField(ref _isDone, value)) OnPropertyChanged(nameof(Badge)); } }

    /// <summary>동그라미 안 글자: 다 했으면 ✔, 아니면 번호</summary>
    public string Badge => _isDone ? "✔" : _number.ToString();

    private string _summary = "";
    /// <summary>제목 아래 작은 글씨 (예: "장비 5대 · 인원 23명")</summary>
    public string Summary { get => _summary; set => SetField(ref _summary, value); }
}
