namespace FormFiller.Core.Models;

/// <summary>
/// 하루치 입력 데이터를 담는 최상위 모델.
/// UI ViewModel → Core 엔진으로 데이터를 전달할 때 이 객체를 사용한다.
/// </summary>
public class DailyInput
{
    /// <summary>작성 대상 날짜</summary>
    public DateTime Date { get; set; }

    /// <summary>날씨 (예: "맑음", "흐림")</summary>
    public string Weather { get; set; } = "";

    /// <summary>금일 오전 작업사항</summary>
    public string TodayMorningWork { get; set; } = "";

    /// <summary>금일 오후 작업사항</summary>
    public string TodayAfternoonWork { get; set; } = "";

    /// <summary>
    /// 금일 작업사항 전체 (오전+오후 합산). 기존 today_work 필드 호환용.
    /// </summary>
    public string TodayWork
    {
        get
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(TodayMorningWork))
                parts.Add(TodayMorningWork);
            if (!string.IsNullOrWhiteSpace(TodayAfternoonWork))
                parts.Add(TodayAfternoonWork);
            return string.Join("\n", parts);
        }
        set
        {
            // 하위 호환: 단순히 오전에 모두 넣음
            TodayMorningWork = value;
            TodayAfternoonWork = "";
        }
    }

    /// <summary>명일 작업사항</summary>
    public string TomorrowWork { get; set; } = "";

    /// <summary>TBM 인원 수</summary>
    public int TbmPersonnel { get; set; }

    /// <summary>장비 투입현황 목록</summary>
    public List<EquipmentEntry> Equipment { get; set; } = [];

    /// <summary>인원 투입현황 목록 (공사팀)</summary>
    public List<PersonnelEntry> Personnel { get; set; } = [];

    /// <summary>수장재팀 별도 인원 목록</summary>
    public List<PersonnelEntry> SujangjaePersonnel { get; set; } = [];

    /// <summary>
    /// 문서 설정에 따라 올바른 인원 목록을 반환한다.
    /// includesSujangjae=true이면 공사인원 + 수장재인원 합산, false이면 공사인원만 반환.
    /// </summary>
    public List<PersonnelEntry> GetPersonnelFor(bool includesSujangjae)
        => includesSujangjae
            ? [.. Personnel, .. SujangjaePersonnel]
            : Personnel;

    /// <summary>TBM 인원 합계 (수장재 포함)</summary>
    public int TotalPersonnelWithSujangjae
        => Personnel.Sum(p => p.TodayCount) + SujangjaePersonnel.Sum(p => p.TodayCount);

    /// <summary>공사 인원 합계 (수장재 제외)</summary>
    public int TotalPersonnelWithoutSujangjae
        => Personnel.Sum(p => p.TodayCount);

    // ── 문서별 직접 입력 값 (input / photo / shape_position 필드) ─────

    /// <summary>
    /// 문서·시트·필드마다 따로 받는 값. 키는 Key(문서, 시트, 필드) 로 만든다.
    /// - input          : 입력한 글자 (단위 제외, 예: "55.3")
    /// - photo          : 사진 파일 경로
    /// - photo 의 측정값 : Key(..., 필드 + ValueSuffixKey) 에 입력한 글자
    /// - shape_position : 고른 위치 이름
    /// </summary>
    public Dictionary<string, string> DocValues { get; set; } = [];

    /// <summary>photo 필드의 짝 값(측정값)을 저장할 때 필드 ID 뒤에 붙이는 표시</summary>
    public const string ValueSuffixKey = "#value";

    public static string Key(string docId, string sheet, string fieldId) => $"{docId}|{sheet}|{fieldId}";

    /// <summary>날짜만 바꾼 복사본 (목록·사전은 원본과 같이 쓴다). 매주 문서를 "그 주 월요일 날짜"로 만들 때 쓴다.</summary>
    public DailyInput WithDate(DateTime date)
    {
        var copy = (DailyInput)MemberwiseClone();   // MemberwiseClone: 필드를 그대로 복사한 얕은 복사본
        copy.Date = date;
        return copy;
    }

    public string GetDocValue(string docId, string sheet, string fieldId)
        => DocValues.TryGetValue(Key(docId, sheet, fieldId), out var v) ? v : "";
}

/// <summary>
/// 장비 한 행의 데이터.
/// CumulativeCount(누계)는 계산 프로퍼티이므로 엑셀 J열(수식 셀)은 건드리지 않는다.
/// </summary>
public class EquipmentEntry
{
    /// <summary>장비명 (예: "B/H 380") — 셀지도에서 행을 찾는 키</summary>
    public string Name { get; set; } = "";

    /// <summary>규격 (예: "39.6t")</summary>
    public string Spec { get; set; } = "";

    /// <summary>전일(대) — 어제 엑셀의 J열(누계)에서 읽어 온 값</summary>
    public int PrevDayCount { get; set; }

    /// <summary>금일(대) — 사용자 입력</summary>
    public int TodayCount { get; set; }

    /// <summary>앞 날짜 파일의 금일(대) — [전일과 같게] 버튼용. 파일에는 쓰지 않는다.</summary>
    public int LastFileTodayCount { get; set; }

    /// <summary>
    /// 누계 = 전일 + 금일.
    /// 엑셀의 J열은 이미 =F+H 수식이 있으므로 여기서는 참조용으로만 사용.
    /// </summary>
    public int CumulativeCount => PrevDayCount + TodayCount;
}

/// <summary>
/// 인원 한 행의 데이터.
/// TotalCumulativeCount(총누계)는 계산 프로퍼티이므로 엑셀 R열(수식 셀)은 건드리지 않는다.
/// </summary>
public class PersonnelEntry
{
    /// <summary>직종 (예: "비계") — 셀지도에서 행을 찾는 키</summary>
    public string Role { get; set; } = "";

    /// <summary>전일누계(명) — 어제 엑셀의 R열(총누계)에서 읽어 온 값</summary>
    public int PrevCumulativeCount { get; set; }

    /// <summary>금일(명) — 사용자 입력</summary>
    public int TodayCount { get; set; }

    /// <summary>앞 날짜 파일의 금일(명) — [전일과 같게] 버튼용. 파일에는 쓰지 않는다.</summary>
    public int LastFileTodayCount { get; set; }

    /// <summary>
    /// 총누계 = 전일누계 + 금일.
    /// 엑셀의 R열은 이미 =N+P 수식이 있으므로 여기서는 참조용으로만 사용.
    /// </summary>
    public int TotalCumulativeCount => PrevCumulativeCount + TodayCount;
}
