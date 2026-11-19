using FluentAssertions;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Core.Models;
using FormFiller.Tests.Helpers;

namespace FormFiller.Tests.Core;

/// <summary>
/// 식 블록(사용자가 직접 만드는 계산)과 "0 일 때 표시"(0 그대로 / 빈칸 / -) 테스트.
/// </summary>
public class BlockFormulaTests : IDisposable
{
    private const string Out = "output", Sheet = "S";
    private static readonly DateTime Mon = new(2026, 11, 16), Tue = new(2026, 11, 17);
    private readonly TempDir _dir = new();
    public void Dispose() => _dir.Dispose();

    private static object Eval(string formula, Dictionary<string, object?>? refs = null) =>
        BlockFormula.Evaluate(formula, r => refs != null && refs.TryGetValue($"{r.Kind}:{r.Arg}", out var v) ? v : null);

    // ── 식 계산 ──────────────────────────────────────────────────────

    [Theory]
    [InlineData("1 + 2 * 3", 7)]
    [InlineData("(1 + 2) * 3", 9)]
    [InlineData("10 / 4", 2.5)]
    [InlineData("-3 + 5", 2)]
    [InlineData("7 / 0", 0)]
    [InlineData("1,200 + 1", 1201)]
    public void 식_사칙연산(string formula, double expected) =>
        Eval(formula).Should().Be((decimal)expected);

    [Fact]
    public void 식_칸과_입력_참조()
    {
        var refs = new Dictionary<string, object?>
        {
            ["어제:H20"] = "12", ["입력:장비"] = 3m, ["칸:B2"] = "5",
        };
        Eval("[어제 H20] + [입력 장비] * [B2]", refs).Should().Be(27m);
        Eval("[어제 h20]", refs).Should().Be("12", "칸 주소는 대문자로");
        Eval("[어제 Z9] + 1", refs).Should().Be(1m, "빈 칸은 0");
    }

    [Fact]
    public void 식_글자_이어붙이기()
    {
        var refs = new Dictionary<string, object?> { ["입력:장비"] = 3m };
        Eval("\"장비 \" & [입력 장비] & \"대\"", refs).Should().Be("장비 3대");
        Eval("\"합계 \" & (1 + 2)").Should().Be("합계 3");
    }

    [Fact]
    public void 식_참조_목록()
    {
        BlockFormula.Refs("[어제 H20] + [입력 장비] & [b2]").Should().Equal(
            new BlockFormula.Ref("어제", "H20"), new BlockFormula.Ref("입력", "장비"), new BlockFormula.Ref("칸", "B2"));
    }

    [Theory]
    [InlineData("1 +", "식이 끝났습니다")]
    [InlineData("(1 + 2", ")")]
    [InlineData("[어제 H20", "]")]
    [InlineData("\"글자", "\"")]
    [InlineData("1 # 2", "#")]
    [InlineData("[어제 ]", "칸")]
    public void 식_틀리면_이유를_알려준다(string formula, string part) =>
        BlockFormula.Check(formula).Should().Contain(part);

    [Fact]
    public void 식_맞으면_Check_는_null() => BlockFormula.Check("[어제 A1] + 1").Should().BeNull();

    // ── 생성: 식 ─────────────────────────────────────────────────────

    private DocumentConfig Doc(params FieldConfig[] fields) => new()
    {
        Id = "f", Name = "f", TemplateFile = "f.xlsx", OutputPattern = "{yy}.{MM}.{dd} f.xlsx", SheetName = Sheet,
        Fields = [.. fields],
    };

    private static FieldConfig Block(string type, string label, Action<FieldConfig> set)
    {
        var f = new FieldConfig { Id = label, Label = label, Type = type };
        set(f);
        return f;
    }

    private static BlockValue B(DailyInput i, string key)
    {
        if (!i.Blocks.TryGetValue(key, out var v)) i.Blocks[key] = v = new BlockValue();
        return v;
    }

    private string Gen(DocumentConfig doc, DateTime date, Action<DailyInput>? edit = null, List<string>? notes = null)
    {
        var input = new DailyInput { Date = date };
        edit?.Invoke(input);
        var problems = new List<string>();
        string path = GenericExcelWriter.Generate(input, doc, _dir.Path, Out, notes: notes, problems: problems);
        problems.Should().BeEmpty();
        return path;
    }

    [Fact]
    public void 생성_식은_어제_파일_칸과_오늘_입력과_오늘_칸으로_계산()
    {
        XlsxBuilder.Create(_dir.File("f.xlsx"), new SheetSpec(Sheet) { ["A1"] = 0, ["B1"] = 10, ["C1"] = 0, ["D1"] = "" });
        var doc = Doc(
            // 식이 앞에 있어도 다른 항목을 다 쓴 뒤 계산한다
            Block(BlockTypes.Formula, "합", f => { f.Cell = "C1"; f.Formula = "[어제 C1] + [A1]"; }),
            Block(BlockTypes.Value, "오늘", f => f.Cell = "A1"),
            Block(BlockTypes.Formula, "글", f => { f.Cell = "D1"; f.Formula = "\"누계 \" & [어제 C1] + [입력 오늘]"; }));

        var notes = new List<string>();
        string mon = Gen(doc, Mon, i => B(i, "오늘").Text = "5", notes);
        notes.Should().Contain(n => n.Contains("어제 파일이 없어"), "첫날은 어제 칸이 0");
        using (var r = new XlsxReader(mon))
        {
            r.Int(Sheet, "C1").Should().Be(5, "어제 없음(0) + 오늘 칸 A1(5)");
            r.Value(Sheet, "D1").Should().Be("누계 5");
        }

        string tue = Gen(doc, Tue, i => B(i, "오늘").Text = "3");
        using var t = new XlsxReader(tue);
        t.Int(Sheet, "C1").Should().Be(8, "어제 C1(5) + 오늘 A1(3)");
        t.Value(Sheet, "D1").Should().Be("누계 8");
    }

    [Fact]
    public void 생성_틀린_식은_쓰지_않고_알린다()
    {
        XlsxBuilder.Create(_dir.File("f.xlsx"), new SheetSpec(Sheet) { ["A1"] = "그대로" });
        var doc = Doc(Block(BlockTypes.Formula, "틀림", f => { f.Cell = "A1"; f.Formula = "1 +"; }));
        var notes = new List<string>();
        using var r = new XlsxReader(Gen(doc, Mon, notes: notes));
        r.Value(Sheet, "A1").Should().Be("그대로");
        notes.Should().Contain(n => n.Contains("틀림") && n.Contains("식"));
    }

    // ── 생성: 0 일 때 표시 ───────────────────────────────────────────

    [Fact]
    public void 영일때_빈칸_또는_대시()
    {
        XlsxBuilder.Create(_dir.File("f.xlsx"), new SheetSpec(Sheet)
        {
            ["A1"] = 9, ["B1"] = 9, ["C1"] = 9, ["D1"] = 9, ["E1"] = 9, ["F1"] = 9,
            ["A3"] = "굴삭기", ["B3"] = 9, ["C3"] = 9, ["D3"] = 9,
        });
        var doc = Doc(
            Block(BlockTypes.Value,   "값",   f => { f.Cell = "A1"; f.ZeroAs = BlockTypes.ZeroDash; }),
            Block(BlockTypes.Counter, "누계", f => { f.Cell = "B1"; f.TotalCell = "C1"; f.ZeroAs = BlockTypes.ZeroBlank; }),
            Block(BlockTypes.Formula, "식",   f => { f.Cell = "D1"; f.Formula = "2 - 2"; f.ZeroAs = BlockTypes.ZeroDash; }),
            Block(BlockTypes.Formula, "그대로", f => { f.Cell = "E1"; f.Formula = "0"; }),
            Block(BlockTypes.Value,   "값2",  f => { f.Cell = "F1"; f.ZeroAs = BlockTypes.ZeroDash; }),
            Block(BlockTypes.Table,   "표",   f =>
            {
                f.StartRow = 3; f.EndRow = 3; f.ZeroAs = BlockTypes.ZeroDash;
                f.Columns = [new() { Role = BlockTypes.RoleName, Col = "A" }, new() { Role = BlockTypes.RolePrev, Col = "B" },
                             new() { Role = BlockTypes.RoleToday, Col = "C" }, new() { Role = BlockTypes.RoleTotal, Col = "D" }];
            }));

        string path = Gen(doc, Mon, i =>
        {
            B(i, "값").Text = "0";
            B(i, "값2").Text = "7";
            B(i, "누계").Prev = 0; B(i, "누계").Today = 0;
            B(i, "표").Rows = [new TableRowValue { Name = "굴삭기", Prev = 4, Today = 0 }];
        });

        using var r = new XlsxReader(path);
        r.Value(Sheet, "A1").Should().Be("-");
        r.Value(Sheet, "B1").Should().Be("");
        r.Value(Sheet, "C1").Should().Be("");
        r.Value(Sheet, "D1").Should().Be("-");
        r.Int(Sheet, "E1").Should().Be(0, "정하지 않으면 0 그대로");
        r.Int(Sheet, "F1").Should().Be(7, "0 이 아니면 숫자");
        r.Int(Sheet, "B3").Should().Be(4);
        r.Value(Sheet, "C3").Should().Be("-");
        r.Int(Sheet, "D3").Should().Be(4);
    }

    [Fact]
    public void 영일때_대시로_쓴_칸도_다음날_0으로_읽어_누계가_이어진다()
    {
        XlsxBuilder.Create(_dir.File("f.xlsx"), new SheetSpec(Sheet) { ["B1"] = 0, ["C1"] = 0 });
        var doc = Doc(Block(BlockTypes.Counter, "누계", f => { f.Cell = "B1"; f.TotalCell = "C1"; f.ZeroAs = BlockTypes.ZeroDash; }));
        Gen(doc, Mon, i => B(i, "누계").Today = 0);

        var prev = GenericExcelWriter.FindPreviousFile(Tue, doc.OutputPattern, _dir.Path, Out)!.Value.Path;
        var init = BlockEngine.InitialValues(doc, BlockEngine.ReadValues(prev, doc));
        init["누계"].Prev.Should().Be(0);

        string tue = Gen(doc, Tue, i => { B(i, "누계").Prev = init["누계"].Prev; B(i, "누계").Today = 2; });
        using var r = new XlsxReader(tue);
        r.Int(Sheet, "C1").Should().Be(2);
    }

    [Fact]
    public void 한글용_글자에도_영일때_표시() =>
        BlockEngine.TextFor(Block(BlockTypes.Formula, "식", f => { f.Formula = "1 - 1"; f.ZeroAs = BlockTypes.ZeroDash; }),
                            new DailyInput { Date = Mon }).Should().Be("-");
}
