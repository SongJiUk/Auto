namespace FormFiller.Core.Config;

/// <summary>
/// config.json 최상위 구조. 문서 목록과 출력 폴더를 보관한다.
/// AppConfigLoader 가 JSON → 이 객체로 역직렬화한다.
/// </summary>
public class AppConfig
{
    /// <summary>생성된 파일을 저장할 폴더 (exe 기준 상대경로)</summary>
    public string OutputFolder { get; set; } = "output";

    /// <summary>처리할 문서 목록. config.json 의 "documents" 배열에 대응.</summary>
    public List<DocumentConfig> Documents { get; set; } = [];
}

/// <summary>
/// 단일 문서(엑셀 파일) 설정.
/// templateFile 에서 읽어 outputPattern 으로 저장하며,
/// fields 배열이 각 셀 처리 방법을 기술한다.
/// </summary>
public class DocumentConfig
{
    /// <summary>문서 식별자 (예: "daily_report", "tbm")</summary>
    public string Id { get; set; } = "";

    /// <summary>사람이 읽을 문서 이름 (예: "공사일보_출력일보")</summary>
    public string Name { get; set; } = "";

    /// <summary>템플릿 파일 경로 (exe 기준 상대경로, 예: "samples/공사일보_출력일보_26.11.04.xlsx")</summary>
    public string TemplateFile { get; set; } = "";

    /// <summary>
    /// 출력 파일명 패턴. {yy}=연도 2자리, {MM}=월, {dd}=일 로 치환.
    /// 예) "공사일보_출력일보_{yy}.{MM}.{dd}.xlsx"
    /// </summary>
    public string OutputPattern { get; set; } = "";

    /// <summary>작성 대상 시트 이름</summary>
    public string SheetName { get; set; } = "";

    /// <summary>이 문서에 채울 필드 목록</summary>
    public List<FieldConfig> Fields { get; set; } = [];

    /// <summary>
    /// 생성 주기. "daily" | "biweekly" | "monthly" | "custom"
    /// </summary>
    public string Frequency { get; set; } = "daily";

    /// <summary>
    /// custom 주기일 때 며칠마다 생성하는지 (예: 7 = 1주마다)
    /// </summary>
    public int FrequencyDays { get; set; } = 1;

    /// <summary>
    /// 수장재팀 인원을 포함해서 계산하는지 여부.
    /// true: 공사인원 + 수장재인원 합산값 사용 (TBM일지 등)
    /// false: 공사인원만 사용 (공사일보, 출력일보 등)
    /// </summary>
    public bool IncludesSujangjae { get; set; } = false;

    /// <summary>
    /// 사진을 삽입할 시트 이름 (예: "사진대지").
    /// 비어 있으면 사진 기능을 사용하지 않는다.
    /// </summary>
    public string PhotoSheet { get; set; } = "";

    /// <summary>
    /// 사진 슬롯별 이름 목록.
    /// 설정 화면에서 각 슬롯에 붙인 이름(예: "장비 사진", "현장 전경")이 저장된다.
    /// </summary>
    public List<PhotoSlotConfig> PhotoSlots { get; set; } = [];
}

/// <summary>
/// 사진대지 시트에서 하나의 이미지 슬롯에 대한 설정.
/// 순서(Index)와 사용자가 붙인 이름(Label)을 저장한다.
/// </summary>
public class PhotoSlotConfig
{
    /// <summary>슬롯 순서 (0-based, 시트 내 이미지 삽입 위치)</summary>
    public int Index { get; set; }

    /// <summary>사용자 지정 이름 (예: "장비 사진"). 비어 있으면 "사진 N"으로 표시.</summary>
    public string Label { get; set; } = "";
}

/// <summary>
/// 단일 필드(셀 또는 테이블) 설정.
/// Type 프로퍼티의 값으로 GenericExcelWriter 가 처리 방식을 결정한다.
/// </summary>
public class FieldConfig
{
    /// <summary>필드 식별자 (예: "date", "weather", "equipment_table")</summary>
    public string Id { get; set; } = "";

    /// <summary>사람이 읽을 필드 이름 (예: "날짜", "날씨")</summary>
    public string Label { get; set; } = "";

    /// <summary>처리 방식 — FieldTypes 상수 중 하나를 사용한다</summary>
    public string Type { get; set; } = "";

    // ── 단일 셀 필드용 ───────────────────────────────────────────────

    /// <summary>단일 셀 주소 (예: "C6"). text·date·number·tbm_personnel 등에서 사용.</summary>
    public string? Cell { get; set; }

    // ── multiline 필드용 ─────────────────────────────────────────────

    /// <summary>멀티라인 텍스트를 나눠 쓸 셀 주소 목록 (예: ["A9","A10",...,"A17"])</summary>
    public List<string>? Cells { get; set; }

    // ── 테이블 필드 공통 ─────────────────────────────────────────────

    /// <summary>테이블 시작 행 (1-based)</summary>
    public int StartRow { get; set; }

    /// <summary>테이블 끝 행 (합계 행 제외)</summary>
    public int EndRow { get; set; }

    /// <summary>장비명 열 — equipment_table 에서 행을 찾는 키</summary>
    public string? NameCol { get; set; }

    /// <summary>전일(대) 열 — equipment_table</summary>
    public string? PrevDayCol { get; set; }

    /// <summary>금일(대/명) 열 — equipment_table 및 personnel_table 공통</summary>
    public string? TodayCol { get; set; }

    /// <summary>직종 열 — personnel_table 에서 행을 찾는 키</summary>
    public string? RoleCol { get; set; }

    /// <summary>전일누계(명) 열 — personnel_table</summary>
    public string? PrevCumulativeCol { get; set; }
}

/// <summary>
/// FieldConfig.Type 에 사용하는 문자열 상수 모음.
/// 엔진(GenericExcelWriter)이 이 값으로 처리 방식을 결정한다.
/// </summary>
public static class FieldTypes
{
    /// <summary>한국어 날짜 (예: "2026년 11월 05일 목요일")</summary>
    public const string DateKorean = "date_korean";

    /// <summary>점(.) 구분 날짜 (예: "2026.11.05")</summary>
    public const string DateSimple = "date_simple";

    /// <summary>일반 텍스트</summary>
    public const string Text = "text";

    /// <summary>정수 숫자</summary>
    public const string Number = "number";

    /// <summary>줄바꿈 텍스트 → 셀 범위(Cells)에 나눠 쓰기</summary>
    public const string Multiline = "multiline";

    /// <summary>장비 투입현황 테이블 (전일·금일 쓰기)</summary>
    public const string EquipmentTable = "equipment_table";

    /// <summary>인원 투입현황 테이블 (전일누계·금일 쓰기)</summary>
    public const string PersonnelTable = "personnel_table";

    /// <summary>금일 대수 > 0 인 장비명을 쉼표로 연결해 단일 셀에 쓰기</summary>
    public const string ActiveEquipmentNames = "active_equipment_names";

    /// <summary>TBM 인원 수 (input.TbmPersonnel)</summary>
    public const string TbmPersonnel = "tbm_personnel";

    /// <summary>참석율 — 계획 = 실시이므로 항상 100</summary>
    public const string AttendanceRate = "attendance_rate";
}
