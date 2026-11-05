namespace FormFiller.App.ViewModels;

/// <summary>
/// 장비 투입현황 한 행 (예: B/H 380 / 39.6t)
/// 전일(대) → 어제 파일에서 읽어 올 값. 1단계 엔진 연결 전까지는 0.
/// 금일(대) → 사용자가 직접 입력.
/// 누계(대) → 전일 + 금일 (자동 계산).
/// </summary>
public class EquipmentEntryViewModel : ViewModelBase
{
    private int _todayCount;

    public string Name { get; init; } = "";   // 장비명: B/H 380
    public string Spec { get; init; } = "";   // 규격: 39.6t

    // 전일: 어제 누계 (엔진 연결 후 채워짐, 지금은 0)
    public int PrevDayCount { get; set; }

    public int TodayCount
    {
        get => _todayCount;
        set
        {
            if (SetField(ref _todayCount, value))
                OnPropertyChanged(nameof(TotalCount)); // 누계도 함께 갱신
        }
    }

    // 누계 = 전일 + 금일 (get만 있고 set이 없음 → 읽기 전용, UI에서 편집 불가)
    public int TotalCount => PrevDayCount + TodayCount;
}
