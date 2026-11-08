using System.Windows.Input;

namespace FormFiller.App.ViewModels;

/// <summary>
/// 장비 투입현황 한 행.
/// 장비명·규격·전일·금일 모두 사용자가 편집할 수 있다.
/// 누계(대)는 전일 + 금일 자동 계산.
/// </summary>
public class EquipmentEntryViewModel : ViewModelBase
{
    private string _name    = "";
    private string _spec    = "";
    private int    _prevDayCount;
    private int    _todayCount;

    public string Name
    {
        get => _name;
        set => SetField(ref _name, value);
    }

    public string Spec
    {
        get => _spec;
        set => SetField(ref _spec, value);
    }

    public int PrevDayCount
    {
        get => _prevDayCount;
        set
        {
            if (SetField(ref _prevDayCount, value))
                OnPropertyChanged(nameof(TotalCount));
        }
    }

    public int TodayCount
    {
        get => _todayCount;
        set
        {
            if (SetField(ref _todayCount, value))
                OnPropertyChanged(nameof(TotalCount));
        }
    }

    // 누계 = 전일 + 금일 (읽기 전용)
    public int TotalCount => PrevDayCount + TodayCount;

    /// <summary>이 행을 목록에서 제거하는 커맨드. 부모 ViewModel이 주입한다.</summary>
    public ICommand RemoveCommand { get; }

    public EquipmentEntryViewModel(Action<EquipmentEntryViewModel>? onRemove = null)
    {
        RemoveCommand = new RelayCommand(_ => onRemove?.Invoke(this));
    }
}
