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

    /// <summary>
    /// 작업사항 항목 선택 팝업에서 보여줄 미리 설정된 작업 목록.
    /// 사용자가 추가·수정·삭제 가능.
    /// </summary>
    public List<string> PresetWorkItems { get; set; } =
    [
        "굴삭기 작업",
        "철거 작업",
        "자재 반출",
        "안전 점검",
        "현장 정리",
        "콘크리트 타설",
        "배관 작업",
        "전기 작업",
        "가설 울타리 설치",
        "폐기물 처리",
    ];

    /// <summary>
    /// Claude AI 채팅 기능에 사용할 Anthropic API 키.
    /// 비어 있으면 AI 기능 비활성화. 절대 git에 올리지 말 것.
    /// </summary>
    public string ClaudeApiKey { get; set; } = "";

    /// <summary>파일 생성이 끝나면 만든 파일을 바로 Excel(또는 한글)로 열지 여부</summary>
    public bool OpenAfterGenerate { get; set; } = true;

    /// <summary>처음 실행할 때 사용법 창을 이미 보여 줬는지 (false 면 시작할 때 자동으로 연다)</summary>
    public bool HelpShown { get; set; } = false;
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

    /// <summary>작성 대상 시트 이름 (SheetNames 가 비어 있을 때 사용)</summary>
    public string SheetName { get; set; } = "";

    /// <summary>
    /// 칸 배치가 같은 시트 여러 개에 같은 설정을 적용할 때의 시트 목록.
    /// (예: 소음측정일지의 "오전초등학교", "다사랑병원", "여성회관")
    /// 비어 있으면 SheetName 한 개만 쓴다. 매일 입력 값(사진·측정값 등)은 시트마다 따로 받는다.
    /// </summary>
    public List<string> SheetNames { get; set; } = [];

    /// <summary>
    /// 사진을 자동으로 가져올 폴더 (예: 구글 드라이브 동기화 폴더).
    /// 이 폴더(또는 시트 이름이 들어간 하위 폴더)에서 촬영 날짜·시각으로 사진을 찾아 넣는다.
    /// </summary>
    public string PhotoFolder { get; set; } = "";

    /// <summary>실제로 채울 시트 목록: SheetNames 가 있으면 그것, 없으면 SheetName 하나.</summary>
    public List<string> TargetSheets() =>
        SheetNames.Count > 0 ? SheetNames : [SheetName];

    /// <summary>이 문서에 채울 필드 목록</summary>
    public List<FieldConfig> Fields { get; set; } = [];

    /// <summary>
    /// 생성 주기. "daily" | "weekly" | "biweekly" | "monthly" | "custom"
    /// </summary>
    public string Frequency { get; set; } = "daily";

    /// <summary>
    /// custom 주기일 때 며칠마다 생성하는지 (예: 7 = 1주마다)
    /// </summary>
    public int FrequencyDays { get; set; } = 1;

    /// <summary>
    /// weekly 주기일 때 만드는 요일 (기본 월요일). config.json 에는 숫자로 저장된다 (0=일요일, 1=월요일 … 6=토요일).
    /// </summary>
    public DayOfWeek WeeklyDay { get; set; } = DayOfWeek.Monday;

    /// <summary>
    /// weekly 주기일 때 파일 날짜를 어떻게 정할지.
    /// false(기본): 만드는 날 날짜 / true: 그 주의 WeeklyDay 날짜 (예: 수요일에 만들어도 월요일 날짜로)
    /// </summary>
    public bool UseWeeklyDayDate { get; set; } = false;

    /// <summary>
    /// 수장재팀 인원을 포함해서 계산하는지 여부.
    /// true: 공사인원 + 수장재인원 합산값 사용 (TBM일지 등)
    /// false: 공사인원만 사용 (공사일보, 출력일보 등)
    /// </summary>
    public bool IncludesSujangjae { get; set; } = false;

    /// <summary>
    /// 파일 저장 방식.
    /// "create": 템플릿 복사 후 새 파일 생성 (기본값, 공사일보 등)
    /// "update": 기존 출력 파일을 직접 수정 (노임대장 등 누적 파일)
    /// </summary>
    public string SaveMode { get; set; } = "create";

    /// <summary>
    /// 앞 날짜에 만든 파일이 있으면 그 파일을 바탕으로 만든다 (기본 켜짐).
    /// 손으로 할 때처럼 "어제 파일 복사 → 바뀐 칸만 고치기" 와 같아서, 지정하지 않은 칸(명단·특기사항·그림 위치 등)은 어제 값이 이어진다.
    /// 앞 파일이 없으면(첫날) 등록한 서식을 쓴다. 한글(hwpx)은 자리 글자가 필요해서 항상 서식을 쓴다.
    /// </summary>
    public bool UsePreviousAsBase { get; set; } = true;

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

    /// <summary>규격 열 — equipment_table (새 장비 줄을 빈 행에 쓸 때 규격도 같이 쓴다)</summary>
    public string? SpecCol { get; set; }

    /// <summary>전일(대) 열 — equipment_table</summary>
    public string? PrevDayCol { get; set; }

    /// <summary>금일(대/명) 열 — equipment_table 및 personnel_table 공통</summary>
    public string? TodayCol { get; set; }

    /// <summary>직종 열 — personnel_table 에서 행을 찾는 키</summary>
    public string? RoleCol { get; set; }

    /// <summary>전일누계(명) 열 — personnel_table</summary>
    public string? PrevCumulativeCol { get; set; }

    // ── input / photo 필드용 ─────────────────────────────────────────

    /// <summary>
    /// 입력 값 뒤에 자동으로 붙일 단위 (예: "dB"). input 타입과 photo 타입의 측정값 칸에 쓴다.
    /// </summary>
    public string? Suffix { get; set; }

    /// <summary>
    /// photo 타입: 사진과 짝이 되는 값 칸 (예: 측정사진 D12 ↔ 측정값 E12). 없으면 사진만 넣는다.
    /// </summary>
    public string? ValueCell { get; set; }

    /// <summary>
    /// photo 타입: 이 사진을 찍는 시각 (예: "10:00").
    /// 사진 폴더에서 불러올 때 촬영 시각이 가장 가까운 사진을 이 칸에 넣는다.
    /// </summary>
    public string? Time { get; set; }

    // ── shape_position 필드용 ────────────────────────────────────────

    /// <summary>함께 옮길 그림·글상자 이름 (예: ["그림 7", "TextBox 10"]) — 화면 표시용</summary>
    public List<string>? ShapeNames { get; set; }

    /// <summary>함께 옮길 그림·글상자의 시트 안 순서 (0부터). 실제로 찾을 때는 이것을 쓴다.</summary>
    public List<int>? ShapeIndexes { get; set; }

    /// <summary>저장해 둔 위치 목록. 매일 이 중 하나를 고른다.</summary>
    public List<ShapePreset>? Presets { get; set; }

    // ── 한글(hwpx) 용 ────────────────────────────────────────────────

    /// <summary>
    /// hwpx 서식에서 바꿀 자리 표시 글자 (예: "{{날짜}}").
    /// 한글에서 서식의 해당 자리에 이 글자를 적어 두면 생성할 때 값으로 바뀐다.
    /// </summary>
    public string? Placeholder { get; set; }
}

/// <summary>
/// 그림·글상자 묶음의 저장된 위치 하나 (예: "정문 앞").
/// </summary>
public class ShapePreset
{
    /// <summary>위치 이름 (예: "정문 앞")</summary>
    public string Name { get; set; } = "";

    /// <summary>묶음 안의 그림·글상자 각각의 위치</summary>
    public List<ShapePlacement> Shapes { get; set; } = [];
}

/// <summary>
/// 그림·글상자 하나의 위치. 엑셀 내부 형식 그대로 "몇 번째 열/행 + 그 칸 안에서 얼마나 떨어졌는지"로 저장한다.
/// 열·행은 0부터 센다. 오프셋 단위는 EMU (1pt = 12,700 EMU).
/// </summary>
public class ShapePlacement
{
    public string Name { get; set; } = "";

    /// <summary>
    /// 시트 안 그림 순서 (0부터, 그림·글상자 전체 기준). -1 이면 이름으로 찾는다.
    /// 같은 모양 시트들은 복사본이라 순서는 같지만 그림 이름은 시트마다 다를 수 있어서 순서를 우선 쓴다.
    /// (예: 장비 그림이 시트1 에선 "그림 21", 시트2 에선 "그림 22")
    /// </summary>
    public int Index { get; set; } = -1;
    public int  FromCol    { get; set; }
    public long FromColOff { get; set; }
    public int  FromRow    { get; set; }
    public long FromRowOff { get; set; }
    public int  ToCol      { get; set; }
    public long ToColOff   { get; set; }
    public int  ToRow      { get; set; }
    public long ToRowOff   { get; set; }
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

    /// <summary>매일 직접 입력하는 값 (시트마다 따로). Suffix 가 있으면 뒤에 붙인다.</summary>
    public const string Input = "input";

    /// <summary>사진 칸 (기존 그림을 새 사진으로 교체). ValueCell 이 있으면 측정값도 함께 받는다.</summary>
    public const string Photo = "photo";

    /// <summary>그림·글상자 묶음의 위치 (저장된 위치 중 선택)</summary>
    public const string ShapePosition = "shape_position";

    /// <summary>매일 입력 화면에서 값을 받아야 하는 타입인지</summary>
    public static bool IsDailyInput(string type) => type is Input or Photo or ShapePosition;
}
