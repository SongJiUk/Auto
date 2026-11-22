using System.Text.Json;
using FluentAssertions;
using FormFiller.Core.Config;
using FormFiller.Core.Models;
using FormFiller.Core.Phone;
using FormFiller.Tests.Helpers;

namespace FormFiller.Tests.Core;

/// <summary>
/// 폰 입력 → 켜 둔 PC 가 자동으로 문서 생성 (2026-11-23).
/// 폰 양식(phone_form.json) 만들기, 폰 입력(input.json) 검사, 어제 누계 위에 얹어 생성, result.json 남기기.
/// </summary>
public class PhoneAutoGeneratorTests : IDisposable
{
    private const string Sheet = "S", Me = "OFFICE-PC";
    private static readonly DateTime Mon = new(2026, 11, 16), Tue = new(2026, 11, 17);
    private static readonly DateTime Now = new(2026, 11, 17, 9, 0, 0);

    private readonly TempDir _dir = new();
    private readonly string _inbox;
    public void Dispose() => _dir.Dispose();

    public PhoneAutoGeneratorTests()
    {
        _inbox = _dir.File("inbox");
        Directory.CreateDirectory(_inbox);
        // 장비 표: 이름 A, 규격 C, 금일 E, 누계 G (서식의 누계 37 / 7)
        XlsxBuilder.Create(_dir.File("t.xlsx"), new SheetSpec(Sheet)
        {
            ["A20"] = "B/H 380", ["C20"] = "39.6t", ["E20"] = 1, ["G20"] = 37,
            ["A21"] = "덤프",    ["C21"] = "15m3",  ["E21"] = 2, ["G21"] = 7,
            ["B5"] = "맑음",
        });
    }

    private static AppConfig Cfg() => new()
    {
        OutputFolder = "output",
        People = [new PersonConfig { Name = "홍길동", Role = "반장" }],
        Documents =
        [
            new DocumentConfig
            {
                Id = "daily", Name = "공사일보", TemplateFile = "t.xlsx", SheetName = Sheet,
                OutputPattern = "{yy}.{MM}.{dd} 공사일보.xlsx",
                Fields =
                [
                    new() { Id = "d", Label = "날짜", Type = BlockTypes.Date, Cell = "A1" },
                    new() { Id = "w", Label = "날씨", Type = BlockTypes.Value, Cell = "B5" },
                    new()
                    {
                        Id = "eq", Label = "장비", Type = BlockTypes.Table, StartRow = 20, EndRow = 23,
                        Columns =
                        [
                            new() { Role = BlockTypes.RoleName, Col = "A" }, new() { Role = BlockTypes.RoleSpec, Col = "C" },
                            new() { Role = BlockTypes.RoleToday, Col = "E" }, new() { Role = BlockTypes.RoleTotal, Col = "G" },
                        ],
                    },
                    new() { Id = "sum", Label = "합계", Type = BlockTypes.Calc, Cell = "B6", CalcKind = BlockTypes.CalcSum, Sources = ["장비"] },
                ],
            },
        ],
    };

    private string Submit(DateTime date, object body, string name = "phone", params string[] photos)
    {
        string folder = Path.Combine(_inbox, $"{date:yyyy-MM-dd}_090000_{name}");
        Directory.CreateDirectory(folder);
        File.WriteAllText(Path.Combine(folder, "input.json"), JsonSerializer.Serialize(body, PhoneDaily.Json));
        foreach (var p in photos) File.WriteAllBytes(Path.Combine(folder, p), [1, 2, 3]);
        return folder;
    }

    private static object Body(DateTime date, object[] rows, string weather = "흐림") => new
    {
        schemaVersion = 2, date = $"{date:yyyy-MM-dd}", device = "android", docs = new[] { "daily" },
        blocks = new object[] { new { key = "날씨", text = weather }, new { key = "장비", rows } },
    };

    private PhoneResult? Run(string folder, DateTime? now = null) =>
        PhoneAutoGenerator.Process(folder, Cfg(), _dir.Path, Me, now ?? Now, null, null);

    // ── 생성 ─────────────────────────────────────────────────────────

    [Fact]
    public void 폰_입력으로_문서를_만들고_결과를_남긴다()
    {
        string folder = Submit(Mon, Body(Mon, [new { name = "B/H 380", spec = "39.6t", today = 2 }]));

        var result = Run(folder);

        result!.Status.Should().Be("done");
        result.Files.Should().Equal("26.11.16 공사일보.xlsx");
        File.Exists(Path.Combine(folder, "result.json")).Should().BeTrue();
        using var r = new XlsxReader(_dir.File(@"output\26.11.16 공사일보.xlsx"));
        r.Value(Sheet, "B5").Should().Be("흐림");
        r.Int(Sheet, "E20").Should().Be(2);
        r.Int(Sheet, "G20").Should().Be(39, "서식 누계 37 + 폰 금일 2");
        r.Int(Sheet, "E21").Should().Be(0, "폰에 없는 줄은 금일 0");
        r.Int(Sheet, "G21").Should().Be(7, "누계는 그대로 이어짐");
    }

    [Fact]
    public void 다음날은_어제_파일의_누계에_더하고_새_장비는_줄을_더한다()
    {
        Run(Submit(Mon, Body(Mon, [new { name = "B/H 380", spec = "39.6t", today = 2 }])));
        Run(Submit(Tue, Body(Tue, [new { name = "B/H 380", spec = "39.6t", today = 3 }, new { name = "B/H 6W", spec = "14.7t", today = 1 }])));

        using var r = new XlsxReader(_dir.File(@"output\26.11.17 공사일보.xlsx"));
        r.Int(Sheet, "G20").Should().Be(42, "39 + 3");
        r.Value(Sheet, "A22").Should().Be("B/H 6W");
        r.Int(Sheet, "E22").Should().Be(1);
    }

    [Fact]
    public void 처리한_폴더와_예전_형식은_다시_하지_않는다()
    {
        string done = Submit(Mon, Body(Mon, []), "a");
        Run(done);
        string old = Submit(Mon, new { schemaVersion = 1, date = "2026-11-16" }, "b");
        string fresh = Submit(Mon, Body(Mon, []), "c");

        PhoneAutoGenerator.FindPending(_inbox).Should().Equal(fresh);
        old.Should().NotBeNull();
    }

    // ── 검사 ─────────────────────────────────────────────────────────

    [Fact]
    public void 숫자가_범위_밖이면_만들지_않고_실패로_남긴다()
    {
        string folder = Submit(Mon, Body(Mon, [new { name = "B/H 380", spec = "", today = 5000 }]));

        var result = Run(folder);

        result!.Status.Should().Be("failed");
        result.Messages.Should().Contain(m => m.Contains("5000"));
        Directory.Exists(_dir.File("output")).Should().BeFalse();
    }

    [Fact]
    public void 사진_파일_이름에_경로가_있으면_거절()
    {
        var body = new
        {
            schemaVersion = 2, date = "2026-11-16", docs = new[] { "daily" },
            slots = new[] { new { doc = "daily", sheet = "사진", index = 1, photo = @"..\..\secret.jpg" } },
        };
        Run(Submit(Mon, body))!.Messages.Should().Contain(m => m.Contains("올바르지 않습니다"));
    }

    [Fact]
    public void 사진이_아직_안_올라왔으면_기다렸다가_오래되면_실패()
    {
        var body = new
        {
            schemaVersion = 2, date = "2026-11-16", docs = new[] { "daily" },
            slots = new[] { new { doc = "daily", sheet = "사진", index = 1, photo = "p1.jpg" } },
        };
        string folder = Submit(Mon, body);
        DateTime created = Directory.GetCreationTime(folder);

        Run(folder, created.AddMinutes(5)).Should().BeNull("드라이브 동기화를 기다린다");
        File.Exists(Path.Combine(folder, "result.json")).Should().BeFalse();
        Run(folder, created.AddMinutes(31))!.Status.Should().Be("failed");
    }

    [Fact]
    public void 다른_PC가_막_맡은_폴더는_건드리지_않는다()
    {
        string folder = Submit(Mon, Body(Mon, []));
        File.WriteAllText(Path.Combine(folder, "claim.txt"), $"HOME-PC|{Now.AddMinutes(-2):o}");
        Run(folder).Should().BeNull();
        Run(folder, Now.AddMinutes(20))!.Status.Should().Be("done", "오래된 맡음은 무시");
    }

    // ── 폰 양식 ──────────────────────────────────────────────────────

    [Fact]
    public void 양식에는_입력할_블록과_어제_장비_사람_목록이_들어간다()
    {
        var cfg = Cfg();
        var initial = PhoneDaily.InitialBlocks(cfg, _dir.Path, Mon);
        var form = PhoneDaily.BuildForm(cfg, Mon, initial, _ => true, _ => [], Me, Now);

        form.Blocks.Select(b => b.Key).Should().Equal("날씨", "장비");   // 날짜·계산은 PC 가 채움
        form.Blocks[1].Rows.Select(r => (r.Name, r.LastToday)).Should().Equal(("B/H 380", 1m), ("덤프", 2m));
        form.People.Should().ContainSingle(p => p.Name == "홍길동");
        form.Docs.Should().ContainSingle(d => d.Id == "daily" && d.Due);
        JsonSerializer.Serialize(form, PhoneDaily.Json).Should().NotContain("ClaudeApiKey").And.NotContain("apiKey");
    }

    [Fact]
    public void 양식_파일은_내용이_같으면_다시_쓰지_않는다()
    {
        var cfg = Cfg();
        var initial = PhoneDaily.InitialBlocks(cfg, _dir.Path, Mon);
        PhoneDaily.WriteForm(_inbox, PhoneDaily.BuildForm(cfg, Mon, initial, _ => true, _ => [], Me, Now)).Should().BeTrue();
        PhoneDaily.WriteForm(_inbox, PhoneDaily.BuildForm(cfg, Mon, initial, _ => true, _ => [], Me, Now.AddMinutes(10))).Should().BeFalse();
    }

    [Fact]
    public void 사진_칸의_측정값은_짝_키로_들어간다()
    {
        var input = new DailyInput();
        var sub = new PhoneSubmission
        {
            DocFields = [new DocFieldValue("noise", "1층", "m10", "55.3", @"C:\x\p.jpg"), new DocFieldValue("noise", "1층", "memo", "비고", null)],
        };
        PhoneSubmissionReader.ApplyTo(input, sub, new Dictionary<string, FieldConfig>(),
                                      new HashSet<(string, string)> { ("noise", "m10") }, []);

        input.GetDocValue("noise", "1층", "m10").Should().Be(@"C:\x\p.jpg");
        input.GetDocValue("noise", "1층", "m10" + DailyInput.ValueSuffixKey).Should().Be("55.3");
        input.GetDocValue("noise", "1층", "memo").Should().Be("비고");
    }
}
