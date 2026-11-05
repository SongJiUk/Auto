using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace FormFiller.App.ViewModels;

/// <summary>
/// 모든 ViewModel의 기반 클래스.
/// INotifyPropertyChanged: 프로퍼티 값이 바뀌었을 때 UI에 알리는 인터페이스.
/// Unity의 [SerializeField]가 Inspector를 갱신하는 것과 비슷한 개념이다.
/// </summary>
public abstract class ViewModelBase : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    /// <summary>
    /// 프로퍼티 setter에서 호출한다. UI가 해당 프로퍼티를 자동으로 다시 읽는다.
    /// [CallerMemberName]: 호출한 프로퍼티 이름을 자동으로 넘겨준다. 직접 쓸 필요 없음.
    /// </summary>
    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    protected bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }
}
