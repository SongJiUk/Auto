namespace FormFiller.Core.Phone;

/// <summary>
/// 폰 앱이 하루치로 보낸 입력. 드라이브 inbox 폴더의 input.json 을 읽어 검증까지 끝난 값이다.
/// 전일·누계는 폰이 모르므로 없다 — PC가 앞 날짜 파일에서 이어받는다.
/// record: 값이 같으면 같은 것으로 보는 읽기 전용 데이터 묶음. `with { ... }` 로 일부만 바꾼 복사본을 만든다.
/// </summary>
public sealed record PhoneInput
{
    public DateTime Date { get; init; }
    public string Device { get; init; } = "";
    public string Weather { get; init; } = "";
    public IReadOnlyList<string> TodayMorningWork { get; init; } = [];
    public IReadOnlyList<string> TodayAfternoonWork { get; init; } = [];
    public IReadOnlyList<string> TomorrowWork { get; init; } = [];
    public IReadOnlyList<PhoneCount> Equipment { get; init; } = [];
    public IReadOnlyList<PhoneCount> Personnel { get; init; } = [];
    public IReadOnlyList<PhoneCount> SujangjaePersonnel { get; init; } = [];
    public IReadOnlyList<PhoneAttendee> Attendees { get; init; } = [];
    public IReadOnlyList<PhonePhoto> Photos { get; init; } = [];
}

/// <summary>장비 한 줄(Name=장비명) 또는 인원 한 줄(Name=직종)의 금일 수</summary>
public sealed record PhoneCount(string Name, string Spec, int TodayCount);

/// <summary>오늘 출근자 한 명. Team 은 PhoneInbox.TEAM_* 중 하나</summary>
public sealed record PhoneAttendee(string Name, string Role, string Team);

/// <summary>사진 한 장. FullPath 는 검증을 통과한 실제 파일 경로 (submission 폴더 안)</summary>
public sealed record PhonePhoto(string File, string Label, string Value, string FullPath);

/// <summary>읽기 결과. Errors 가 있으면 Input 은 null</summary>
public sealed record PhoneReadResult(PhoneInput? Input, IReadOnlyList<string> Errors);
