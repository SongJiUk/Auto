using FluentAssertions;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Tests.Helpers;

namespace FormFiller.Tests.Core;

/// <summary>머리글 자동 인식. 칸 배치는 올려 주신 공사일보·출력일보와 같은 모양.</summary>
public class HeaderDetectorTests : IDisposable
{
    private readonly TempDir _dir = new();
    public void Dispose() => _dir.Dispose();

    private Func<int, int, string?> Sheet(SheetSpec spec)
    {
        string path = XlsxBuilder.Create(_dir.File($"{Guid.NewGuid():N}.xlsx"), spec);
        var doc = CellHelper.OpenReadOnly(path);
        return HeaderDetector.FromFile(CellHelper.FindWorksheetPart(doc.WorkbookPart!, spec.Name)!);
    }

    private SheetSpec DailyReport() => new("공사일보")
    {
        ["A18"] = "[ 장 비   투 입 현 황 ]", ["L18"] = "[ 인 원   투 입 현 황 ]",
        ["A19"] = "장비명", ["D19"] = "규격", ["F19"] = "전일(대)", ["H19"] = "금일(대)", ["J19"] = "누계(대)",
        ["L19"] = "직종", ["N19"] = "전일누계(명)", ["P19"] = "금일(명)", ["R19"] = "총 누계(명)", ["T19"] = "비  고",
        ["A20"] = "B/H 0.2", ["L20"] = "관리자",
    };

    [Fact]
    public void 장비_표_머리글()
    {
        // A20:K32 를 드래그 (A=1 … K=11)
        var cols = HeaderDetector.Detect(BlockTypes.Table, 20, 1, 11, Sheet(DailyReport()));
        cols.Should().BeEquivalentTo(new Dictionary<string, string>
            { ["name"] = "A", ["spec"] = "D", ["prev"] = "F", ["today"] = "H", ["total"] = "J" });
    }

    [Fact]
    public void 인원_표_머리글_전일누계는_전일_총누계는_누계()
    {
        // L20:T32 (L=12 … T=20)
        var cols = HeaderDetector.Detect(BlockTypes.Table, 20, 12, 20, Sheet(DailyReport()));
        cols.Should().BeEquivalentTo(new Dictionary<string, string>
            { ["name"] = "L", ["prev"] = "N", ["today"] = "P", ["total"] = "R" });
    }

    [Fact]
    public void 한_열만_드래그해도_옆_열에서_찾는다()
    {
        var cols = HeaderDetector.Detect(BlockTypes.Table, 20, 1, 1, Sheet(DailyReport()));
        cols["name"].Should().Be("A");
        cols["today"].Should().Be("H", "왼쪽 표(장비)의 금일이 먼저");
    }

    [Fact]
    public void 출력일보_명단_왼쪽_절반만()
    {
        var spec = new SheetSpec("출력일보")
        {
            ["A1"] = "일일출력일보", ["G2"] = "2026년 11월 17일 화요일",
            ["A3"] = "NO", ["B3"] = "직 종", ["C3"] = "성 명", ["D3"] = "작업내용",
            ["E3"] = "NO", ["F3"] = "직 종", ["G3"] = "성 명", ["H3"] = "작업내용",
            ["A4"] = 1,
        };
        var get = Sheet(spec);
        HeaderDetector.Detect(BlockTypes.Roster, 4, 1, 4, get).Should().BeEquivalentTo(new Dictionary<string, string>
            { ["job"] = "B", ["person"] = "C", ["task"] = "D" });
        HeaderDetector.Detect(BlockTypes.Roster, 4, 5, 8, get).Should().BeEquivalentTo(new Dictionary<string, string>
            { ["job"] = "F", ["person"] = "G", ["task"] = "H" }, "오른쪽 절반(21~40번)을 드래그하면 오른쪽 열");
    }

    [Fact]
    public void 머리글이_없으면_빈_결과()
    {
        var cols = HeaderDetector.Detect(BlockTypes.Table, 5, 1, 6, Sheet(new SheetSpec("S") { ["A5"] = "x" }));
        cols.Should().BeEmpty();
    }

    [Fact]
    public void 금일_열이_없으면_전일_열이_매일_입력_칸()
    {
        // 폐기물 반출: 구분 · 규격 · 전일반출(대) · 누계(대) · 총 반출수량 — 금일 열이 없다.
        // 사용자 규칙 (2026-11-22): 누계 = 어제 누계 + 오늘 적는 전일반출 → 전일반출은 "금일"(매일 입력) 역할
        var spec = new SheetSpec("S")
        {
            ["A34"] = "구분", ["C34"] = "규격", ["E34"] = "전일반출(대)", ["G34"] = "누계(대)", ["I34"] = "총 반출수량",
            ["A35"] = "덤프",
        };
        var cols = HeaderDetector.Detect(BlockTypes.Table, 35, 1, 10, Sheet(spec));
        cols.Should().BeEquivalentTo(new Dictionary<string, string>
            { ["name"] = "A", ["spec"] = "C", ["today"] = "E", ["total"] = "G" });
    }

    [Fact]
    public void 철거_전일철거동_누계동도_같은_규칙()
    {
        var spec = new SheetSpec("S")
        {
            ["L35"] = "공종", ["N35"] = "총건물동수", ["P35"] = "전일철거동", ["R35"] = "누계동", ["T35"] = "잔여동", ["V35"] = "공정율",
            ["L36"] = "철거동",
        };
        var cols = HeaderDetector.Detect(BlockTypes.Table, 36, 12, 23, Sheet(spec));
        cols.Should().BeEquivalentTo(new Dictionary<string, string>
            { ["name"] = "L", ["today"] = "P", ["total"] = "R" });
    }

    [Fact]
    public void 머리글_한_줄_위에_제목이_있어도_가까운_머리글_줄을_쓴다()
    {
        var spec = new SheetSpec("S")
        {
            ["A1"] = "전일 금일 현황표",                            // 제목 (한 칸뿐이라 머리글로 안 봄)
            ["A2"] = "구분", ["B2"] = "전일", ["C2"] = "금일", ["D2"] = "합계",
        };
        HeaderDetector.Detect(BlockTypes.Table, 3, 1, 4, Sheet(spec)).Should().BeEquivalentTo(new Dictionary<string, string>
            { ["name"] = "A", ["prev"] = "B", ["today"] = "C", ["total"] = "D" });
    }

    [Theory]
    [InlineData(BlockTypes.Table, "전일누계(명)", "prev")]
    [InlineData(BlockTypes.Table, "총 누계(명)", "total")]
    [InlineData(BlockTypes.Table, "금 일", "today")]
    [InlineData(BlockTypes.Table, "비  고", null)]
    [InlineData(BlockTypes.Roster, "성 명", "person")]
    [InlineData(BlockTypes.Roster, "NO", null)]
    public void 머리글_글자_하나(string kind, string header, string? role) =>
        HeaderDetector.RoleFor(kind, header).Should().Be(role);
}
