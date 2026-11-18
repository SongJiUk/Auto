using FluentAssertions;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Core.Models;
using FormFiller.Tests.Helpers;

namespace FormFiller.Tests.Core;

/// <summary>문서 주기(매일·매주·2주·매월·N일) 계산 테스트. 2026-11-16 = 월요일.</summary>
public class DocumentScheduleTests
{
    private static readonly DateTime Mon = new(2026, 11, 16), Tue = new(2026, 11, 17), Wed = new(2026, 11, 18),
                                     Sun = new(2026, 11, 22), NextMon = new(2026, 11, 23);

    private static DocumentConfig Doc(string freq, DayOfWeek day = DayOfWeek.Monday, int days = 0, bool dayDate = false,
                                      string pattern = "{yy}.{MM}.{dd} 작업계획서.xlsx") =>
        new() { Frequency = freq, WeeklyDay = day, FrequencyDays = days, UseWeeklyDayDate = dayDate, OutputPattern = pattern };

    // ── 주의 시작·요일 ───────────────────────────────────────────────

    [Fact]
    public void 주는_월요일에_시작한다()
    {
        DocumentSchedule.WeekStart(Wed).Should().Be(Mon);
        DocumentSchedule.WeekStart(Sun).Should().Be(Mon, "일요일은 그 주의 마지막 날");
        DocumentSchedule.WeekStart(Mon).Should().Be(Mon);
        DocumentSchedule.DayInWeek(Sun, DayOfWeek.Friday).Should().Be(new DateTime(2026, 11, 20));
    }

    // ── 매주 ─────────────────────────────────────────────────────────

    [Fact]
    public void 매주_월요일_한_번도_안_만들었으면_월요일부터_차례()
    {
        var d = Doc(DocumentSchedule.Weekly);
        DocumentSchedule.IsDue(d, Mon, null).Should().BeTrue();
        DocumentSchedule.IsDue(d, Wed, null).Should().BeTrue("월요일에 못 만들었으면 지나도 계속 차례");
    }

    [Fact]
    public void 매주_이번_주에_만들었으면_다음_주까지_아님()
    {
        var d = Doc(DocumentSchedule.Weekly);
        DocumentSchedule.IsDue(d, Wed, last: Mon).Should().BeFalse();
        DocumentSchedule.IsDue(d, Sun, last: Tue).Should().BeFalse();
        DocumentSchedule.NextDue(d, Wed, last: Mon).Should().Be(NextMon);
        DocumentSchedule.IsDue(d, NextMon, last: Mon).Should().BeTrue();
    }

    [Fact]
    public void 매주_지난주에_만든_것은_이번_주와_상관없다()
    {
        var d = Doc(DocumentSchedule.Weekly);
        DocumentSchedule.IsDue(d, Mon, last: new DateTime(2026, 11, 13)).Should().BeTrue("지난주 금요일에 만든 것");
    }

    [Fact]
    public void 매주_금요일_문서는_금요일_전에는_아직()
    {
        var d = Doc(DocumentSchedule.Weekly, DayOfWeek.Friday);
        DocumentSchedule.IsDue(d, Wed, null).Should().BeFalse();
        DocumentSchedule.NextDue(d, Wed, null).Should().Be(new DateTime(2026, 11, 20));
        DocumentSchedule.StatusText(d, Wed, null).Should().Be("이번 주 금요일(11/20)부터");
    }

    [Fact]
    public void 매주_상태_글자()
    {
        var d = Doc(DocumentSchedule.Weekly);
        DocumentSchedule.StatusText(d, Wed, last: Mon).Should().Be("이번 주 완료 (11/16)");
        DocumentSchedule.StatusText(d, Wed, null).Should().Be("이번 주 아직 — 오늘 만들 차례");
        DocumentSchedule.Label(d).Should().Be("매주 월요일");
    }

    [Fact]
    public void 매주_파일_날짜는_설정에_따라_만드는_날_또는_정한_요일()
    {
        DocumentSchedule.DocumentDate(Doc(DocumentSchedule.Weekly), Wed).Should().Be(Wed, "기본은 만드는 날");
        DocumentSchedule.DocumentDate(Doc(DocumentSchedule.Weekly, dayDate: true), Wed).Should().Be(Mon);
        DocumentSchedule.DocumentDate(Doc(DocumentSchedule.Daily, dayDate: true), Wed).Should().Be(Wed, "매일 문서엔 적용 안 함");
    }

    // ── 그 밖의 주기 ─────────────────────────────────────────────────

    [Fact]
    public void 매일은_항상_차례()
    {
        DocumentSchedule.IsDue(Doc(DocumentSchedule.Daily), Wed, last: Wed).Should().BeTrue();
        DocumentSchedule.IsDaily(Doc(DocumentSchedule.Daily)).Should().BeTrue();
        DocumentSchedule.IsDaily(Doc(DocumentSchedule.Weekly)).Should().BeFalse();
    }

    [Fact]
    public void 이주마다는_14일_지나면()
    {
        var d = Doc(DocumentSchedule.Biweekly);
        DocumentSchedule.IsDue(d, Wed, last: new DateTime(2026, 11, 5)).Should().BeFalse("13일");
        DocumentSchedule.IsDue(d, Wed, last: new DateTime(2026, 11, 4)).Should().BeTrue("14일");
        DocumentSchedule.NextDue(d, Wed, last: new DateTime(2026, 11, 5)).Should().Be(new DateTime(2026, 11, 19));
    }

    [Fact]
    public void 매월은_달이_바뀌면()
    {
        var d = Doc(DocumentSchedule.Monthly);
        DocumentSchedule.IsDue(d, Wed, last: new DateTime(2026, 11, 2)).Should().BeFalse();
        DocumentSchedule.IsDue(d, Wed, last: new DateTime(2026, 10, 30)).Should().BeTrue();
        DocumentSchedule.NextDue(d, Wed, last: new DateTime(2026, 11, 2)).Should().Be(new DateTime(2026, 12, 1));
        DocumentSchedule.IsDue(d, new DateTime(2027, 11, 3), last: new DateTime(2026, 11, 2)).Should().BeTrue("해가 바뀜");
    }

    [Fact]
    public void N일마다()
    {
        var d = Doc(DocumentSchedule.Custom, days: 10);
        DocumentSchedule.IsDue(d, Wed, last: new DateTime(2026, 11, 9)).Should().BeFalse();
        DocumentSchedule.IsDue(d, Wed, last: new DateTime(2026, 11, 8)).Should().BeTrue();
        DocumentSchedule.Label(d).Should().Be("10일마다");
    }

    // ── 마지막으로 만든 날짜 (파일 이름에서) ──────────────────────────

    [Theory]
    [InlineData("{yy}.{MM}.{dd} 공사일보(참마루건설).xlsx", "26.11.14 공사일보(참마루건설).xlsx", "2026-11-14")]
    [InlineData("공사일보_출력일보_{yy}.{MM}.{dd}.xlsx", "공사일보_출력일보_26.11.05.xlsx", "2026-11-05")]
    [InlineData("TBM_{yy}{MM}{dd}.xlsx", "TBM_261118.xlsx", "2026-11-18")]
    [InlineData("{yy}.{MM}.{dd} 공사일보(참마루건설).xlsx", "26.11.14 공사일보(참마루건설) - 복사본.xlsx", null)]
    [InlineData("{yy}.{MM}.{dd} 공사일보.xlsx", "26.11.14 TBM.xlsx", null)]
    [InlineData("{yy}.{MM}.{dd} 공사일보.xlsx", "26.13.40 공사일보.xlsx", null)]   // 없는 날짜
    [InlineData("작업계획서.xlsx", "작업계획서.xlsx", null)]                        // 날짜 없는 패턴
    public void 파일_이름에서_날짜_읽기(string pattern, string file, string? expected) =>
        DocumentSchedule.DateFromFileName(pattern, file).Should().Be(expected == null ? null : DateTime.Parse(expected));

    [Fact]
    public void 날짜로_시작하는_저장_이름도_마지막_날짜를_찾는다()
    {
        // 예전 화면 코드는 패턴의 '{' 앞 글자로 찾아서, 날짜로 시작하는 이름은 못 찾았다 (= 주기 문서가 매일 체크됨)
        var files = new[] { "26.11.09 계획서.xlsx", "26.11.16 계획서.xlsx", "26.11.17 공사일보.xlsx", "메모.txt" };
        DocumentSchedule.LastGenerated("{yy}.{MM}.{dd} 계획서.xlsx", files).Should().Be(Mon);
    }

    [Fact]
    public void 고른_날짜_뒤의_파일은_보지_않는다()
    {
        var files = new[] { "26.11.09 계획서.xlsx", "26.11.23 계획서.xlsx" };
        DocumentSchedule.LastGenerated("{yy}.{MM}.{dd} 계획서.xlsx", files, notAfter: Wed)
            .Should().Be(new DateTime(2026, 11, 9));
    }

    [Fact]
    public void 출력_폴더에서_찾는다()
    {
        using var dir = new TempDir();
        File.WriteAllText(dir.File("26.11.16 작업계획서.xlsx"), "");
        File.WriteAllText(dir.File("26.11.02 작업계획서.xlsx"), "");
        DocumentSchedule.LastGenerated(Doc(DocumentSchedule.Weekly), dir.Path).Should().Be(Mon);
        DocumentSchedule.LastGenerated(Doc(DocumentSchedule.Weekly), Path.Combine(dir.Path, "없음")).Should().BeNull();
    }

    // ── 엔진: 매주 문서를 정한 요일 날짜로 만들기 ───────────────────

    [Fact]
    public void 정한_요일_날짜로_설정하면_수요일에_만들어도_월요일_파일()
    {
        using var dir = new TempDir();
        XlsxBuilder.Create(dir.File("plan.xlsx"), new SheetSpec("계획") { ["B2"] = "옛 날짜" });
        var cfg = new DocumentConfig
        {
            Id = "plan", Name = "작업계획서", TemplateFile = "plan.xlsx", SheetName = "계획",
            OutputPattern = "{yy}.{MM}.{dd} 작업계획서.xlsx",
            Frequency = DocumentSchedule.Weekly, WeeklyDay = DayOfWeek.Monday, UseWeeklyDayDate = true,
            Fields = [new() { Id = "date", Type = FieldTypes.DateKorean, Cell = "B2" }],
        };
        var input = new DailyInput { Date = Wed };

        string path = GenericExcelWriter.Generate(input, cfg, dir.Path, "output");

        Path.GetFileName(path).Should().Be("26.11.16 작업계획서.xlsx");
        GenericExcelWriter.OutputPathFor(cfg, Wed, dir.Path, "output").Should().Be(path);
        using var r = new XlsxReader(path);
        r.Value("계획", "B2").Should().Be("2026년 11월 16일 월요일");
        input.Date.Should().Be(Wed, "화면 입력값(원본)은 바꾸지 않는다");
    }
}
