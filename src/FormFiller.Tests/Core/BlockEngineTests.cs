using FluentAssertions;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Core.Models;
using FormFiller.Tests.Helpers;

namespace FormFiller.Tests.Core;

/// <summary>
/// 블록(사용자가 규칙을 고르는 범용 항목) 테스트.
/// 화면이 하는 순서 그대로: 앞 날짜 파일 읽기(ReadValues) → 처음 값(InitialValues) → 입력 → 생성(Generate)
/// </summary>
public class BlockEngineTests : IDisposable
{
    private const string Out = "output", Sheet = "S";
    private static readonly DateTime Mon = new(2026, 11, 16), Tue = new(2026, 11, 17);
    private readonly TempDir _dir = new();
    public void Dispose() => _dir.Dispose();

    private static FieldConfig Block(string type, string label, Action<FieldConfig>? set = null)
    {
        var f = new FieldConfig { Id = label, Label = label, Type = type };
        set?.Invoke(f);
        return f;
    }

    private DocumentConfig Doc(string template, params FieldConfig[] fields) => new()
    {
        Id = Path.GetFileNameWithoutExtension(template), Name = template, TemplateFile = template,
        OutputPattern = "{yy}.{MM}.{dd} " + Path.GetFileNameWithoutExtension(template) + ".xlsx",
        SheetName = Sheet, Fields = [.. fields],
    };

    /// <summary>화면 흐름: 앞 날짜 파일(없으면 서식)에서 처음 값을 만들고, edit 로 입력한 뒤 생성</summary>
    private (string Path, List<string> Notes) Day(DocumentConfig doc, DateTime date, Action<DailyInput>? edit = null)
    {
        var prev = GenericExcelWriter.FindPreviousFile(date, doc.OutputPattern, _dir.Path, Out);
        var read = BlockEngine.ReadValues(prev?.Path ?? _dir.File(doc.TemplateFile), doc);
        var input = new DailyInput { Date = date, Blocks = BlockEngine.InitialValues(doc, read) };
        edit?.Invoke(input);
        var notes = new List<string>();
        var problems = new List<string>();
        string path = GenericExcelWriter.Generate(input, doc, _dir.Path, Out, notes: notes, problems: problems);
        problems.Should().BeEmpty("자동 검증 통과");
        return (path, notes);
    }

    private static XlsxReader Open(string path) => new(path);

    // ── value ────────────────────────────────────────────────────────

    [Fact]
    public void 값_매일입력_어제값유지_고정값()
    {
        XlsxBuilder.Create(_dir.File("v.xlsx"), new SheetSpec(Sheet) { ["A1"] = "맑음", ["A2"] = "손으로 쓴 값", ["A3"] = "옛 값", ["A4"] = 1 });
        var doc = Doc("v.xlsx",
            Block(BlockTypes.Value, "날씨", f => f.Cell = "A1"),
            Block(BlockTypes.Value, "비고", f => { f.Cell = "A2"; f.ValueSource = BlockTypes.SourceKeep; }),
            Block(BlockTypes.Value, "참석율", f => { f.Cell = "A3"; f.ValueSource = BlockTypes.SourceFixed; f.FixedValue = "100"; }),
            Block(BlockTypes.Value, "인원", f => f.Cell = "A4"));

        var (path, _) = Day(doc, Mon, i => { i.Blocks["날씨"].Text = "흐림"; i.Blocks["인원"].Text = "15"; });

        using var r = Open(path);
        r.Value(Sheet, "A1").Should().Be("흐림");
        r.Value(Sheet, "A2").Should().Be("손으로 쓴 값", "어제 값 그대로");
        r.Value(Sheet, "A3").Should().Be("100");
        r.RawCell(Sheet, "A3")!.DataType.Should().BeNull("숫자로 쓴다");
        r.Int(Sheet, "A4").Should().Be(15);
    }

    [Fact]
    public void 값_처음값은_어제_파일의_글자()
    {
        XlsxBuilder.Create(_dir.File("v.xlsx"), new SheetSpec(Sheet) { ["A1"] = "55.3dB" });
        var doc = Doc("v.xlsx", Block(BlockTypes.Value, "측정", f => { f.Cell = "A1"; f.Suffix = "dB"; }));
        var init = BlockEngine.InitialValues(doc, BlockEngine.ReadValues(_dir.File("v.xlsx"), doc));
        init["측정"].Text.Should().Be("55.3", "단위는 떼고 보여 준다");
    }

    // ── date ─────────────────────────────────────────────────────────

    [Theory]
    [InlineData("{yyyy}년 {MM}월 {dd}일 {요일}", "2026년 11월 16일 월요일")]
    [InlineData("{yyyy}.{MM}.{dd}", "2026.11.16")]
    [InlineData("{yy}.{MM}.{dd}", "26.11.16")]
    [InlineData("{yyyy}년 {M}월 {d}일", "2026년 11월 16일")]
    [InlineData("{yyyy}.{MM}.{dd}({요})", "2026.11.16(월)")]
    [InlineData(null, "2026년 11월 16일 월요일")]
    public void 날짜_형식(string? format, string expected) =>
        BlockTypes.FormatDate(format, Mon).Should().Be(expected);

    [Fact]
    public void 날짜_한_자리_월일()
    {
        BlockTypes.FormatDate("{M}/{d}", new DateTime(2026, 3, 5)).Should().Be("3/5");
    }

    // ── lines ────────────────────────────────────────────────────────

    [Fact]
    public void 여러줄_앞글자_붙이고_남는칸_비우고_넘치면_알림()
    {
        XlsxBuilder.Create(_dir.File("l.xlsx"), new SheetSpec(Sheet) { ["A1"] = "- 옛1", ["A2"] = "- 옛2", ["A3"] = "- 옛3" });
        var doc = Doc("l.xlsx", Block(BlockTypes.Lines, "작업", f => { f.Cells = ["A1", "A2"]; f.Prefix = "- "; }));

        var (path, notes) = Day(doc, Mon, i => i.Blocks["작업"].Lines = ["철거", "- 반출", "  ", "-", "살수"]);

        using var r = Open(path);
        r.Value(Sheet, "A1").Should().Be("- 철거");
        r.Value(Sheet, "A2").Should().Be("- 반출", "이미 붙어 있으면 다시 붙이지 않는다");
        r.Value(Sheet, "A3").Should().Be("- 옛3", "블록 칸이 아님");
        notes.Should().Contain(n => n.Contains("3줄 중 칸이 2개"));
    }

    [Fact]
    public void 여러줄_금일은_어제의_명일에서_시작_없으면_어제_금일()
    {
        XlsxBuilder.Create(_dir.File("w.xlsx"), new SheetSpec(Sheet)
            { ["A1"] = "- 어제 금일", ["B1"] = "- 어제 명일1", ["B2"] = "- 어제 명일2" });
        var today = Block(BlockTypes.Lines, "금일", f => { f.Cells = ["A1", "A2"]; f.StartFrom = "명일"; });
        var tomorrow = Block(BlockTypes.Lines, "명일", f => { f.Cells = ["B1", "B2"]; f.StartFrom = "same"; });
        var doc = Doc("w.xlsx", today, tomorrow);

        var init = BlockEngine.InitialValues(doc, BlockEngine.ReadValues(_dir.File("w.xlsx"), doc));
        init["금일"].Lines.Should().Equal("- 어제 명일1", "- 어제 명일2");
        init["명일"].Lines.Should().Equal("- 어제 명일1", "- 어제 명일2");

        // 명일이 비어 있으면 어제 금일
        XlsxBuilder.Create(_dir.File("w2.xlsx"), new SheetSpec(Sheet) { ["A1"] = "- 어제 금일" });
        var doc2 = Doc("w2.xlsx", today, tomorrow);
        BlockEngine.InitialValues(doc2, BlockEngine.ReadValues(_dir.File("w2.xlsx"), doc2))["금일"].Lines
            .Should().Equal("- 어제 금일");
    }

    // ── counter ──────────────────────────────────────────────────────

    /// <summary>올려 주신 공사일보의 폐기물 반출·공정율 모양</summary>
    private DocumentConfig CounterDoc()
    {
        XlsxBuilder.Create(_dir.File("c.xlsx"), new SheetSpec(Sheet)
        {
            ["E36"] = 21, ["G36"] = 21, ["I36"] = new F("G36*15", "315"),      // 덤프: 오늘, 누계, 총 반출수량(수식)
            ["P36"] = 0, ["R36"] = 62, ["N36"] = 147, ["T36"] = new F("N36-R36", "85"),  // 공정율: 전일철거, 누계동
        });
        return Doc("c.xlsx",
            Block(BlockTypes.Counter, "덤프", f => { f.Cell = "E36"; f.TotalCell = "G36"; }),
            Block(BlockTypes.Counter, "철거동", f => { f.Cell = "P36"; f.TotalCell = "R36"; }));
    }

    [Fact]
    public void 누계_어제누계_더하기_오늘_이틀_연속()
    {
        var doc = CounterDoc();
        Day(doc, Mon, i => { i.Blocks["덤프"].Today = 5; i.Blocks["철거동"].Today = 2; });
        var (tue, _) = Day(doc, Tue, i => { i.Blocks["덤프"].Today = 3; });

        using var r = Open(tue);
        r.Int(Sheet, "E36").Should().Be(3);
        r.Int(Sheet, "G36").Should().Be(29, "21 + 5 + 3");
        r.Int(Sheet, "P36").Should().Be(0);
        r.Int(Sheet, "R36").Should().Be(64, "62 + 2 + 0");
        r.RawCell(Sheet, "I36")!.CellFormula.Should().NotBeNull("총 반출수량 수식 유지");
    }

    [Fact]
    public void 누계_칸이_수식이면_전일칸_더하기_오늘로_이어받는다()
    {
        XlsxBuilder.Create(_dir.File("c2.xlsx"), new SheetSpec(Sheet) { ["A1"] = 10, ["B1"] = 2, ["C1"] = new F("A1+B1", "12") });
        var doc = Doc("c2.xlsx", Block(BlockTypes.Counter, "n", f => { f.PrevCell = "A1"; f.Cell = "B1"; f.TotalCell = "C1"; }));
        Day(doc, Mon, i => i.Blocks["n"].Today = 4);
        var (tue, _) = Day(doc, Tue, i => i.Blocks["n"].Today = 1);

        using var r = Open(tue);
        r.Int(Sheet, "A1").Should().Be(16, "12 + 4 (낡은 수식 캐시 12 가 아니라 전일+오늘)");
        r.Int(Sheet, "B1").Should().Be(1);
        r.RawCell(Sheet, "C1")!.CellFormula.Should().NotBeNull();
    }

    // ── table ────────────────────────────────────────────────────────

    private DocumentConfig TableDoc(bool withSpec = true, bool withTotalFormula = true)
    {
        var s = new SheetSpec(Sheet)
        {
            ["A20"] = "B/H 380", ["D20"] = "39.6t", ["F20"] = 36, ["H20"] = 1,
            ["A21"] = "덤프",    ["D21"] = "15m3",  ["F21"] = 5,  ["H21"] = 2,
        };
        foreach (int row in new[] { 20, 21, 22, 23 })
            s[$"J{row}"] = withTotalFormula ? new F($"F{row}+H{row}", "0") : 0;
        XlsxBuilder.Create(_dir.File("t.xlsx"), s);
        var cols = new List<ColumnRole>
        {
            new() { Role = BlockTypes.RoleName, Col = "A" }, new() { Role = BlockTypes.RolePrev, Col = "F" },
            new() { Role = BlockTypes.RoleToday, Col = "H" }, new() { Role = BlockTypes.RoleTotal, Col = "J" },
        };
        if (withSpec) cols.Add(new() { Role = BlockTypes.RoleSpec, Col = "D" });
        return Doc("t.xlsx", Block(BlockTypes.Table, "장비", f => { f.StartRow = 20; f.EndRow = 23; f.Columns = cols; }));
    }

    [Fact]
    public void 표_이틀_연속_누계_이월_새줄_지운줄()
    {
        var doc = TableDoc();
        Day(doc, Mon, i =>
        {
            i.Blocks["장비"].Rows[0].Today = 2;                                     // B/H 380
            i.Blocks["장비"].Rows.RemoveAt(1);                                      // 덤프 지움
            i.Blocks["장비"].Rows.Add(new TableRowValue { Name = "B/H 6W", Spec = "14.7t", Today = 1 });
        });
        var (tue, _) = Day(doc, Tue, i => i.Blocks["장비"].Rows.Single(r => r.Name == "B/H 6W").Today = 3);

        using var r = Open(tue);
        r.Int(Sheet, "F20").Should().Be(39, "36+1 → 37, +2");
        r.Int(Sheet, "H20").Should().Be(0);
        r.Int(Sheet, "F21").Should().Be(7, "지운 덤프는 5+2 이월, 금일 0");
        r.Value(Sheet, "A22").Should().Be("B/H 6W");
        r.Value(Sheet, "D22").Should().Be("14.7t");
        r.Int(Sheet, "F22").Should().Be(1);
        r.Int(Sheet, "H22").Should().Be(3);
        r.RawCell(Sheet, "J22")!.CellFormula.Should().NotBeNull();
    }

    [Fact]
    public void 표_누계가_숫자칸이면_써준다()
    {
        var doc = TableDoc(withTotalFormula: false);
        var (path, _) = Day(doc, Mon, i => i.Blocks["장비"].Rows[0].Today = 4);
        using var r = Open(path);
        r.Int(Sheet, "J20").Should().Be(41, "전일 37 + 금일 4");
    }

    [Fact]
    public void 표_처음값은_어제누계와_어제금일()
    {
        var doc = TableDoc();
        var init = BlockEngine.InitialValues(doc, BlockEngine.ReadValues(_dir.File("t.xlsx"), doc));
        init["장비"].Rows.Select(r => (r.Name, r.Spec, r.Prev, r.Today, r.LastToday))
            .Should().Equal(("B/H 380", "39.6t", 37m, 0m, 1m), ("덤프", "15m3", 7m, 0m, 2m));
    }

    // ── roster + 명단으로 인원 세기 ──────────────────────────────────

    private DocumentConfig RosterDoc()
    {
        XlsxBuilder.Create(_dir.File("r.xlsx"),
            new SheetSpec("공사일보")
            {
                ["L20"] = "관리자", ["N20"] = 233, ["P20"] = 3, ["R20"] = new F("N20+P20", "236"),
                ["L21"] = "장비",   ["N21"] = 174, ["P21"] = 10, ["R21"] = new F("N21+P21", "184"),
                ["L22"] = "비계",   ["N22"] = 302,
            },
            new SheetSpec("출력일보")
            {
                ["A4"] = 1, ["B4"] = "소장", ["C4"] = "가", ["D4"] = "현장감독",
                ["A5"] = 2, ["B5"] = "장비", ["C5"] = "나", ["D5"] = "잔재물정리",
                ["A6"] = 3, ["B6"] = "장비", ["C6"] = "다", ["D6"] = "살수",
                ["A7"] = 4,
            });
        var table = Block(BlockTypes.Table, "인원", f =>
        {
            f.StartRow = 20; f.EndRow = 23;
            f.Columns = [new() { Role = BlockTypes.RoleName, Col = "L" }, new() { Role = BlockTypes.RolePrev, Col = "N" },
                         new() { Role = BlockTypes.RoleToday, Col = "P" }, new() { Role = BlockTypes.RoleTotal, Col = "R" }];
            f.RosterKey = "출근자";
            f.RoleMap = new() { ["관리자"] = ["소장", "차장", "대리"] };
        });
        var roster = Block(BlockTypes.Roster, "출근자", f =>
        {
            f.StartRow = 4; f.EndRow = 7;
            f.Columns = [new() { Role = BlockTypes.RoleJob, Col = "B" }, new() { Role = BlockTypes.RolePerson, Col = "C" },
                         new() { Role = BlockTypes.RoleTask, Col = "D" }];
        });
        return new DocumentConfig
        {
            Id = "daily", Name = "공사일보", TemplateFile = "r.xlsx", OutputPattern = "{yy}.{MM}.{dd} 공사일보.xlsx",
            SheetName = "공사일보",
            Fields = [table, roster],
        };
    }

    [Fact]
    public void 명단을_고르면_출력일보에_쓰고_인원표_금일을_직종별로_센다()
    {
        // 한 파일 안 두 시트: 인원 표는 공사일보 시트(적용할 시트), 명단은 출력일보 시트(Sheet 로 따로 지정)
        var doc = RosterDoc();
        doc.Fields[1].Sheet = "출력일보";

        var read = BlockEngine.ReadValues(_dir.File("r.xlsx"), doc);
        var input = new DailyInput { Date = Mon, Blocks = BlockEngine.InitialValues(doc, read) };
        input.Blocks["출근자"].People.Should().HaveCount(3, "어제(서식) 명단이 처음 값");
        input.Blocks["인원"].Rows.Should().HaveCount(3);

        // 오늘: 소장 1, 대리 1, 장비 1, 목수(표에 없음) 1
        input.Blocks["출근자"].People =
        [
            new() { Job = "소장", Person = "가", Task = "현장감독" }, new() { Job = "대리", Person = "라", Task = "현장관리" },
            new() { Job = "장비", Person = "나", Task = "잔재물정리" }, new() { Job = "목수", Person = "바", Task = "" },
        ];
        var notes = new List<string>();
        BlockEngine.ApplyRosterCounts(input, doc.Fields, notes);
        input.Blocks["인원"].Rows.Select(r => (r.Name, r.Today)).Should().Equal(("관리자", 2m), ("장비", 1m), ("비계", 0m));
        notes.Should().ContainSingle(n => n.Contains("목수"));

        var problems = new List<string>();
        string path = GenericExcelWriter.Generate(input, doc, _dir.Path, Out, problems: problems);
        problems.Should().BeEmpty();

        using var r = Open(path);
        r.Int("공사일보", "N20").Should().Be(236, "233 + 3");
        r.Int("공사일보", "P20").Should().Be(2);
        r.Int("공사일보", "P21").Should().Be(1);
        r.Value("출력일보", "B5").Should().Be("대리");
        r.Value("출력일보", "C5").Should().Be("라");
        r.Value("출력일보", "C7").Should().Be("바");
        r.Value("출력일보", "D7").Should().BeEmpty("작업내용 비움");
        r.Int("출력일보", "A7").Should().Be(4, "번호 열은 그대로");
        r.Value("공사일보", "B5").Should().BeNull("명단은 출력일보 시트에만");
    }

    [Fact]
    public void 명단_사람이_줄어들면_남는_줄은_비운다_넘치면_알림()
    {
        XlsxBuilder.Create(_dir.File("ro.xlsx"), new SheetSpec(Sheet)
            { ["B1"] = "장비", ["C1"] = "가", ["B2"] = "장비", ["C2"] = "나", ["D2"] = "살수" });
        var doc = Doc("ro.xlsx", Block(BlockTypes.Roster, "명단", f =>
        {
            f.StartRow = 1; f.EndRow = 2;
            f.Columns = [new() { Role = BlockTypes.RoleJob, Col = "B" }, new() { Role = BlockTypes.RolePerson, Col = "C" },
                         new() { Role = BlockTypes.RoleTask, Col = "D" }];
        }));

        var (one, _) = Day(doc, Mon, i => i.Blocks["명단"].People = [new() { Job = "소장", Person = "다" }]);
        using (var r = Open(one))
        {
            r.Value(Sheet, "C1").Should().Be("다");
            r.Value(Sheet, "B2").Should().BeEmpty();
            r.Value(Sheet, "C2").Should().BeEmpty();
            r.Value(Sheet, "D2").Should().BeEmpty();
        }

        var (_, notes) = Day(doc, Tue, i => i.Blocks["명단"].People =
            [new() { Person = "가" }, new() { Person = "나" }, new() { Person = "다" }]);
        notes.Should().Contain(n => n.Contains("3명 중 줄이 2개"));
    }

    // ── calc ─────────────────────────────────────────────────────────

    [Fact]
    public void 계산_합계_이름나열_인원수()
    {
        var input = new DailyInput { Date = Mon };
        input.Blocks["인원"] = new() { Rows = [new() { Name = "관리자", Today = 3 }, new() { Name = "비계", Today = 8 }] };
        input.Blocks["수장재 인원"] = new() { Text = "4" };
        input.Blocks["덤프"] = new() { Today = 2 };
        input.Blocks["장비"] = new() { Rows = [new() { Name = "B/H 380", Today = 1 }, new() { Name = "덤프", Today = 0 }, new() { Name = "세륜기", Today = 2 }] };
        input.Blocks["출근자"] = new() { People = [new(), new(), new()] };

        string Calc(string kind, params string[] src) =>
            BlockEngine.Calc(new FieldConfig { Type = BlockTypes.Calc, CalcKind = kind, Sources = [.. src] }, input);

        Calc(BlockTypes.CalcSum, "인원", "수장재 인원").Should().Be("15");
        Calc(BlockTypes.CalcSum, "인원", "덤프", "없는 이름").Should().Be("13");
        Calc(BlockTypes.CalcNames, "장비").Should().Be("B/H 380, 세륜기");
        Calc(BlockTypes.CalcCount, "출근자").Should().Be("3");
    }

    // ── 여러 문서가 같은 이름 블록을 함께 쓴다 ──────────────────────

    [Fact]
    public void 같은_이름_블록은_두_문서에_같은_값()
    {
        XlsxBuilder.Create(_dir.File("a.xlsx"), new SheetSpec(Sheet) { ["A1"] = "" });
        XlsxBuilder.Create(_dir.File("b.xlsx"), new SheetSpec(Sheet) { ["C3"] = "" });
        var a = Doc("a.xlsx", Block(BlockTypes.Value, "날씨", f => f.Cell = "A1"));
        var b = Doc("b.xlsx", Block(BlockTypes.Value, "날씨 칸", f => { f.Cell = "C3"; f.Key = "날씨"; }));
        var input = new DailyInput { Date = Mon, Blocks = { ["날씨"] = new BlockValue { Text = "비" } } };

        string pa = GenericExcelWriter.Generate(input, a, _dir.Path, Out);
        string pb = GenericExcelWriter.Generate(input, b, _dir.Path, Out);

        using (var r = Open(pa)) r.Value(Sheet, "A1").Should().Be("비");
        using (var r = Open(pb)) r.Value(Sheet, "C3").Should().Be("비");
    }

    // ── 한글 ─────────────────────────────────────────────────────────

    [Fact]
    public void 한글_자리글자에도_블록값()
    {
        string tpl = _dir.File("plan.hwpx");
        using (var z = System.IO.Compression.ZipFile.Open(tpl, System.IO.Compression.ZipArchiveMode.Create))
        {
            using (var w = new StreamWriter(z.CreateEntry("mimetype", System.IO.Compression.CompressionLevel.NoCompression).Open())) w.Write("application/hwp+zip");
            using (var w = new StreamWriter(z.CreateEntry("Contents/section0.xml").Open())) w.Write("<p>{{날짜}} {{작업}} {{인원}}</p>");
        }
        var doc = new DocumentConfig
        {
            Id = "plan", Name = "계획", TemplateFile = "plan.hwpx", OutputPattern = "{dd} 계획.hwpx",
            Fields =
            [
                Block(BlockTypes.Date, "날짜", f => f.DateFormat = "{yyyy}.{MM}.{dd}"),
                Block(BlockTypes.Lines, "작업", f => f.Prefix = "- "),
                Block(BlockTypes.Calc, "인원", f => { f.CalcKind = BlockTypes.CalcSum; f.Sources = ["n"]; }),
            ],
        };
        var input = new DailyInput { Date = Mon, Blocks = { ["작업"] = new() { Lines = ["철거"] }, ["n"] = new() { Text = "7" } } };

        string path = GenericExcelWriter.Generate(input, doc, _dir.Path, Out);

        using var zip = System.IO.Compression.ZipFile.OpenRead(path);
        using var rd = new StreamReader(zip.GetEntry("Contents/section0.xml")!.Open());
        rd.ReadToEnd().Should().Be("<p>2026.11.16 - 철거 7</p>");
    }
}
