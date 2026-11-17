using FluentAssertions;
using FormFiller.Core.Phone;

namespace FormFiller.Tests.Core;

/// <summary>폰이 드라이브 inbox 폴더에 남긴 입력(input.json + 사진)을 읽고 검증하는 테스트</summary>
public sealed class PhoneInboxTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "ff_inbox_" + Guid.NewGuid().ToString("N"));

    public PhoneInboxTests() => Directory.CreateDirectory(_root);

    public void Dispose()
    {
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
    }

    private string MakeSubmission(string folderName, string json, params string[] photoFiles)
    {
        string dir = Path.Combine(_root, folderName);
        Directory.CreateDirectory(dir);
        File.WriteAllText(Path.Combine(dir, PhoneInbox.INPUT_FILE_NAME), json);
        foreach (var f in photoFiles) File.WriteAllBytes(Path.Combine(dir, f), [0xFF, 0xD8, 0xFF]);
        return dir;
    }

    private const string VALID_JSON = """
    {
      "schemaVersion": 1,
      "date": "2026-11-18",
      "device": "갤럭시",
      "weather": " 맑음 ",
      "todayMorningWork": ["- 3층 내부 철거", "  "],
      "todayAfternoonWork": [],
      "tomorrowWork": ["- 4층 철거"],
      "equipment": [{ "name": "B/H 380", "spec": "39.6t", "todayCount": 2 }],
      "personnel": [{ "role": "비계", "todayCount": 3 }],
      "sujangjaePersonnel": [{ "role": "수장재", "todayCount": 4 }],
      "attendees": [{ "name": "가나다", "role": "비계", "team": "공사" }],
      "photos": [{ "file": "p01.jpg", "label": "장비", "value": "55.3" }]
    }
    """;

    [Fact]
    public void Read_ValidSubmission_ReturnsTrimmedInput()
    {
        string dir = MakeSubmission("2026-11-18_074000_갤럭시", VALID_JSON, "p01.jpg");

        var result = PhoneInbox.Read(dir);

        result.Errors.Should().BeEmpty();
        var input = result.Input!;
        input.Date.Should().Be(new DateTime(2026, 11, 18));
        input.Weather.Should().Be("맑음");
        input.TodayMorningWork.Should().Equal("- 3층 내부 철거");   // 빈 줄은 버림
        input.Equipment.Should().ContainSingle().Which.TodayCount.Should().Be(2);
        input.SujangjaePersonnel.Should().ContainSingle().Which.Name.Should().Be("수장재");   // 인원 줄은 Name 에 직종
        input.Attendees.Should().ContainSingle().Which.Team.Should().Be("공사");
        input.Photos.Should().ContainSingle().Which.FullPath.Should().Be(Path.Combine(dir, "p01.jpg"));
    }

    [Fact]
    public void Read_MissingInputFile_ReturnsError()
    {
        string dir = Path.Combine(_root, "2026-11-18_080000_x");
        Directory.CreateDirectory(dir);

        var result = PhoneInbox.Read(dir);

        result.Input.Should().BeNull();
        result.Errors.Should().ContainSingle();
    }

    [Fact]
    public void Read_BrokenJson_ReturnsErrorWithoutThrowing()
    {
        string dir = MakeSubmission("2026-11-18_080000_x", "{ not json");

        var result = PhoneInbox.Read(dir);

        result.Input.Should().BeNull();
        result.Errors.Should().ContainSingle();
    }

    [Theory]
    [InlineData("""{ "schemaVersion": 2, "date": "2026-11-18" }""")]   // 모르는 형식 버전
    [InlineData("""{ "schemaVersion": 1, "date": "11/18" }""")]        // 날짜 형식 틀림
    [InlineData("""{ "schemaVersion": 1, "date": "2026-11-19" }""")]   // 폴더 날짜와 다름
    [InlineData("""{ "schemaVersion": 1, "date": "2026-11-18", "equipment": [{ "name": "B/H", "todayCount": -1 }] }""")]
    [InlineData("""{ "schemaVersion": 1, "date": "2026-11-18", "personnel": [{ "role": "비계", "todayCount": 1000 }] }""")]
    [InlineData("""{ "schemaVersion": 1, "date": "2026-11-18", "attendees": [{ "name": "가", "team": "외부" }] }""")]
    public void Read_InvalidValues_ReturnsError(string json)
    {
        string dir = MakeSubmission("2026-11-18_080000_x", json);

        var result = PhoneInbox.Read(dir);

        result.Input.Should().BeNull();
        result.Errors.Should().NotBeEmpty();
    }

    private static string PhotoJson(string file) =>
        $$"""{ "schemaVersion": 1, "date": "2026-11-18", "photos": [{ "file": {{System.Text.Json.JsonSerializer.Serialize(file)}}, "label": "장비" }] }""";

    [Theory]
    [InlineData("p01.exe")]
    [InlineData("없는사진.jpg")]
    public void Read_BadPhotoFile_ReturnsError(string file)
    {
        string dir = MakeSubmission("2026-11-18_080000_x", PhotoJson(file), "p01.exe");

        var result = PhoneInbox.Read(dir);

        result.Input.Should().BeNull();
        result.Errors.Should().NotBeEmpty();
    }

    // 실제로 존재하는 폴더 밖 파일을 가리켜도 거절해야 한다 (파일이 없어서 실패하는 것과 구분)
    [Theory]
    [InlineData("../outside.jpg")]
    [InlineData("..\\outside.jpg")]
    [InlineData("sub\\inner.jpg")]
    [InlineData("ABSOLUTE")]
    public void Read_PhotoPathOutsideFileName_ReturnsError(string file)
    {
        File.WriteAllBytes(Path.Combine(_root, "outside.jpg"), [0xFF]);
        string dir = Path.Combine(_root, "2026-11-18_080000_x");
        Directory.CreateDirectory(Path.Combine(dir, "sub"));
        File.WriteAllBytes(Path.Combine(dir, "sub", "inner.jpg"), [0xFF]);
        if (file == "ABSOLUTE") file = Path.Combine(_root, "outside.jpg");
        MakeSubmission("2026-11-18_080000_x", PhotoJson(file));

        var result = PhoneInbox.Read(dir);

        result.Input.Should().BeNull();
        result.Errors.Should().ContainSingle().Which.Should().Contain("올바르지 않습니다");
    }

    [Fact]
    public void FindSubmissions_ReturnsOnlyThatDate_OldestFirst()
    {
        MakeSubmission("2026-11-18_090000_b", "{}");
        MakeSubmission("2026-11-18_074000_a", "{}");
        MakeSubmission("2026-11-17_180000_a", "{}");
        Directory.CreateDirectory(Path.Combine(_root, "2026-11-18_100000_partial"));   // input.json 아직 동기화 안 됨

        var found = PhoneInbox.FindSubmissions(_root, new DateTime(2026, 11, 18));

        found.Select(Path.GetFileName).Should().Equal("2026-11-18_074000_a", "2026-11-18_090000_b");
    }

    [Fact]
    public void FindSubmissions_MissingInboxFolder_ReturnsEmpty()
    {
        PhoneInbox.FindSubmissions(Path.Combine(_root, "없음"), new DateTime(2026, 11, 18))
                  .Should().BeEmpty();
    }
}
