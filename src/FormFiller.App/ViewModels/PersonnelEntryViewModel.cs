using System.Windows.Input;

namespace FormFiller.App.ViewModels;

/// <summary>
/// 인원 투입현황 한 행.
/// 직종·전일누계·금일 모두 사용자가 편집할 수 있다.
/// 총누계는 전일누계 + 금일 자동 계산.
/// </summary>
public class PersonnelEntryViewModel : ViewModelBase
{
    private string _role               = "";
    private int    _prevCumulativeCount;
    private int    _todayCount;

    public string Role
    {
        get => _role;
        set => SetField(ref _role, value);
    }

    public int PrevCumulativeCount
    {
        get => _prevCumulativeCount;
        set
        {
            if (SetField(ref _prevCumulativeCount, value))
                OnPropertyChanged(nameof(TotalCumulativeCount));
        }
    }

    public int TodayCount
    {
        get => _todayCount;
        set
        {
            if (SetField(ref _todayCount, value))
                OnPropertyChanged(nameof(TotalCumulativeCount));
        }
    }

    // 총누계 = 전일누계 + 금일 (읽기 전용)
    public int TotalCumulativeCount => PrevCumulativeCount + TodayCount;

    /// <summary>이 행을 목록에서 제거하는 커맨드. 부모 ViewModel이 주입한다.</summary>
    public ICommand RemoveCommand { get; }

    public PersonnelEntryViewModel(Action<PersonnelEntryViewModel>? onRemove = null)
    {
        RemoveCommand = new RelayCommand(_ => onRemove?.Invoke(this));
    }
}
