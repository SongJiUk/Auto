namespace FormFiller.App.ViewModels;

/// <summary>
/// 인원 투입현황 한 행 (예: 비계 / 전일누계 5 / 금일 3 / 총누계 8)
/// </summary>
public class PersonnelEntryViewModel : ViewModelBase
{
    private int _todayCount;

    public string Role { get; init; } = ""; // 직종: 관리자, 벌목, 비계 …

    // 전일누계: 어제 총누계 (엔진 연결 후 채워짐, 지금은 0)
    public int PrevCumulativeCount { get; set; }

    public int TodayCount
    {
        get => _todayCount;
        set
        {
            if (SetField(ref _todayCount, value))
                OnPropertyChanged(nameof(TotalCumulativeCount));
        }
    }

    // 총누계 = 전일누계 + 금일
    public int TotalCumulativeCount => PrevCumulativeCount + TodayCount;
}
