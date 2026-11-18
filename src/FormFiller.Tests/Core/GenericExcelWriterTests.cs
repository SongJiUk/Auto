using System.Security.Cryptography;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Core.Models;
using FormFiller.Tests.Helpers;

namespace FormFiller.Tests.Core;

/// <summary>
/// 공사일보 엔진(GenericExcelWriter) 테스트. 가짜 서식(DailyReportSample)으로 실제 흐름을 따라 한다:
///   앞 날짜 파일 찾기 → 전일 값 읽기(LoadPreviousDay) → 생성(Generate) → 결과 파일 확인
/// </summary>
public class GenericExcelWriterTests : IDisposable
{
    private const string Out = "output";
    private static readonly DateTime Thu = new(2026, 11, 5);   // 목
    private static readonly DateTime Fri = new(2026, 11, 6);   // 금
    private static readonly DateTime Mon = new(2026, 11, 9);   // 월

    private readonly TempDir _dir = new();
    private readonly string _template;

    public GenericExcelWriterTests()
    {
        _template = DailyReportSample.CreateTemplate(_dir.File(DailyReportSample.TemplateName));
    }

    public void Dispose() => _dir.Dispose();

    /// <summary>화면이 하는 일과 같게: 앞 날짜 파일(없으면 서식)에서 전일 값을 읽은 뒤 생성</summary>
    private string MakeDay(DailyInput input, DocumentConfig? cfg = null, List<string>? notes = null)
    {
        cfg ??= DailyReportSample.Config();
        var prev = GenericExcelWriter.FindPreviousFile(input.Date, cfg.OutputPattern, _dir.Path, Out);
        GenericExcelWriter.LoadPreviousDay(prev?.Path ?? _template, input, cfg);
        return GenericExcelWriter.Generate(input, cfg, _dir.Path, Out, notes: notes);
    }

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));

    // ── 원본 보호 ────────────────────────────────────────────────────

    [Fact]
    public void 서식_원본은_바뀌지_않는다()
    {
        string before = Hash(_template);
        MakeDay(DailyReportSample.Input(Thu));
        MakeDay(DailyReportSample.Input(Fri));
        Hash(_template).Should().Be(before);
    }

    [Fact]
    public void 앞_날짜_파일도_바뀌지_않는다()
    {
        string thu = MakeDay(DailyReportSample.Input(Thu));
        string before = Hash(thu);
        MakeDay(DailyReportSample.Input(Fri));
        Hash(thu).Should().Be(before);
    }

    // ── 파일 이름·날짜 ───────────────────────────────────────────────

    [Fact]
    public void 저장_이름은_패턴대로()
    {
        string path = MakeDay(DailyReportSample.Input(Thu));
        Path.GetFileName(path).Should().Be("26.11.05 공사일보.xlsx");
        Path.GetDirectoryName(path).Should().Be(Path.Combine(_dir.Path, Out));
    }

    [Theory]
    [InlineData(2026, 11, 4, "2026년 11월 04일 수요일")]
    [InlineData(2026, 11, 8, "2026년 11월 08일 일요일")]
    [InlineData(2026, 12, 31, "2026년 12월 31일 목요일")]
    public void 한국어_날짜_형식(int y, int m, int d, string expected) =>
        GenericExcelWriter.FormatKoreanDate(new DateTime(y, m, d)).Should().Be(expected);

    [Fact]
    public void 날짜_날씨_칸은_공유_문자열로_남는다()
    {
        string path = MakeDay(DailyReportSample.Input(Thu));
        using var r = new XlsxReader(path);
        r.Value(DailyReportSample.Sheet, "C6").Should().Be("2026년 11월 05일 목요일");
        r.Value(DailyReportSample.Sheet, "L6").Should().Be("흐림");
        // 서식이 공유 문자열(t="s")이었으면 그대로 공유 문자열로 쓴다
        r.RawCell(DailyReportSample.Sheet, "C6")!.DataType!.Value.Should().Be(CellValues.SharedString);
    }

    // ── 작업사항 (여러 줄) ───────────────────────────────────────────

    [Fact]
    public void 작업사항은_줄마다_칸에_나눠_쓰고_남는_칸은_비운다()
    {
        string path = MakeDay(DailyReportSample.Input(Thu));
        using var r = new XlsxReader(path);
        r.Value(DailyReportSample.Sheet, "A9").Should().Be("철거 작업");
        r.Value(DailyReportSample.Sheet, "A10").Should().Be("자재 반출");
        r.Value(DailyReportSample.Sheet, "A11").Should().BeEmpty("어제 3번째 줄이 남으면 안 된다");
        r.Value(DailyReportSample.Sheet, "L9").Should().Be("현장 정리");
    }

    // ── 수식·재계산 ──────────────────────────────────────────────────

    [Fact]
    public void 누계_합계_수식은_그대로_남는다()
    {
        string path = MakeDay(DailyReportSample.Input(Thu));
        using var r = new XlsxReader(path);
        foreach (var cell in new[] { "J20", "J21", "J23", "F25", "H25", "R20", "R21", "P25" })
            r.RawCell(DailyReportSample.Sheet, cell)!.CellFormula.Should().NotBeNull($"{cell} 은 수식 칸");
    }

    [Fact]
    public void 저장하면_열_때_다시_계산하도록_표시한다()
    {
        string path = MakeDay(DailyReportSample.Input(Thu));
        using var r = new XlsxReader(path);
        r.Workbook.Workbook.GetFirstChild<CalculationProperties>()!.FullCalculationOnLoad!.Value.Should().BeTrue();
    }

    [Fact]
    public void 수식_칸에는_날짜를_쓰지_않는다()
    {
        // 두 번째 시트 날짜가 첫 시트를 참조하는 경우 (소음측정일지 등)
        var s1 = new SheetSpec("시트1") { ["C6"] = "옛 날짜" };
        var s2 = new SheetSpec("시트2") { ["C6"] = new F("시트1!C6", "옛 날짜") };
        XlsxBuilder.Create(_dir.File("두시트.xlsx"), s1, s2);
        var cfg = new DocumentConfig
        {
            Id = "two", Name = "두시트", TemplateFile = "두시트.xlsx", OutputPattern = "{yy}{MM}{dd} 두시트.xlsx",
            SheetNames = ["시트1", "시트2"],
            Fields = [new() { Id = "date", Type = FieldTypes.DateSimple, Cell = "C6" }],
        };

        string path = GenericExcelWriter.Generate(new DailyInput { Date = Thu }, cfg, _dir.Path, Out);

        using var r = new XlsxReader(path);
        r.Value("시트1", "C6").Should().Be("2026.11.05");
        r.RawCell("시트2", "C6")!.CellFormula!.Text.Should().Be("시트1!C6");
    }

    [Fact]
    public void 만든_파일은_스키마_검사를_통과한다()
    {
        string path = MakeDay(DailyReportSample.Input(Thu));
        using var r = new XlsxReader(path);
        r.ValidationErrors().Should().BeEmpty();
    }

    [Fact]
    public void 재계산_표시는_통합문서_요소_순서를_지킨다()
    {
        // calcPr 이 없는 서식 + 뒤에 와야 하는 요소(extLst)가 있을 때, calcPr 을 맨 끝에 붙이면 스키마 위반
        string path = _dir.File("calc.xlsx");
        XlsxBuilder.Create(path, new SheetSpec("S") { ["A1"] = "x" });
        using (var d = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, true))
        {
            d.WorkbookPart!.Workbook.Append(new WorkbookExtensionList());
            d.WorkbookPart.Workbook.Save();
        }
        var cfg = new DocumentConfig
        {
            Id = "c", Name = "c", TemplateFile = "calc.xlsx", OutputPattern = "{dd} c.xlsx", SheetName = "S",
            Fields = [new() { Id = "date", Type = FieldTypes.DateSimple, Cell = "A1" }],
        };
        string outPath = GenericExcelWriter.Generate(new DailyInput { Date = Thu }, cfg, _dir.Path, Out);

        using var r = new XlsxReader(outPath);
        r.ValidationErrors().Should().BeEmpty();
    }

    // ── 장비 표 ──────────────────────────────────────────────────────

    [Fact]
    public void 첫날_장비_전일은_서식의_전일_더하기_금일()
    {
        string path = MakeDay(DailyReportSample.Input(Thu, bh: 2, dump: 0, wash: 1));
        using var r = new XlsxReader(path);
        // 서식: B/H 전일 10 + 금일 1 = 11 → 오늘 전일 11, 금일 2
        r.Int(DailyReportSample.Sheet, "F20").Should().Be(11);
        r.Int(DailyReportSample.Sheet, "H20").Should().Be(2);
        r.Int(DailyReportSample.Sheet, "F21").Should().Be(7);
        r.Int(DailyReportSample.Sheet, "H21").Should().Be(0);
    }

    [Fact]
    public void 누계가_숫자_칸이면_전일_더하기_금일을_써_준다()
    {
        string path = MakeDay(DailyReportSample.Input(Thu, wash: 2));
        using var r = new XlsxReader(path);
        // 세륜기: 전일 3+1=4, 금일 2 → 누계(J22, 숫자 칸) 6
        r.Int(DailyReportSample.Sheet, "F22").Should().Be(4);
        r.Int(DailyReportSample.Sheet, "J22").Should().Be(6);
    }

    [Fact]
    public void 이틀_연속_만들면_누계가_이어진다()
    {
        MakeDay(DailyReportSample.Input(Thu, bh: 2, dump: 3));   // B/H 전일 11, 금일 2 / 덤프 전일 7, 금일 3
        string fri = MakeDay(DailyReportSample.Input(Fri, bh: 4, dump: 1));

        using var r = new XlsxReader(fri);
        r.Int(DailyReportSample.Sheet, "F20").Should().Be(13, "목요일 전일 11 + 금일 2");
        r.Int(DailyReportSample.Sheet, "H20").Should().Be(4);
        r.Int(DailyReportSample.Sheet, "F21").Should().Be(10, "목요일 전일 7 + 금일 3");
        r.Int(DailyReportSample.Sheet, "H21").Should().Be(1);
    }

    [Fact]
    public void 주말을_건너뛰고_금요일_파일에서_이어받는다()
    {
        MakeDay(DailyReportSample.Input(Thu, bh: 2));
        MakeDay(DailyReportSample.Input(Fri, bh: 4));
        string mon = MakeDay(DailyReportSample.Input(Mon, bh: 1));

        GenericExcelWriter.FindPreviousFile(Mon, DailyReportSample.Pattern, _dir.Path, Out)!.Value.Date
            .Should().Be(Fri);
        using var r = new XlsxReader(mon);
        r.Int(DailyReportSample.Sheet, "F20").Should().Be(17, "11 + 2 + 4");
    }

    [Fact]
    public void 프로그램이_만든_파일의_낡은_수식_캐시값을_믿지_않는다()
    {
        // 프로그램이 만든 파일은 누계 수식의 저장된 값이 서식 원본 숫자(11) 그대로다.
        // 전일 값은 반드시 "전일 칸 + 금일 칸" 으로 계산해야 한다.
        string thu = MakeDay(DailyReportSample.Input(Thu, bh: 5));
        using (var r = new XlsxReader(thu))
            r.Int(DailyReportSample.Sheet, "J20").Should().Be(11, "캐시는 갱신되지 않음 (엑셀이 열 때 다시 계산)");

        var input = DailyReportSample.Input(Fri);
        GenericExcelWriter.LoadPreviousDay(thu, input, DailyReportSample.Config());
        input.Equipment[0].PrevDayCount.Should().Be(16, "전일 11 + 금일 5");
        input.Equipment[0].LastFileTodayCount.Should().Be(5);
    }

    [Fact]
    public void 화면에서_지운_장비는_금일_0_전일은_이월()
    {
        var input = DailyReportSample.Input(Thu);
        input.Equipment.RemoveAll(e => e.Name == "덤프");
        string path = MakeDay(input);

        using var r = new XlsxReader(path);
        r.Value(DailyReportSample.Sheet, "A21").Should().Be("덤프", "줄은 남긴다");
        r.Int(DailyReportSample.Sheet, "F21").Should().Be(7, "서식 전일 5 + 금일 2");
        r.Int(DailyReportSample.Sheet, "H21").Should().Be(0);
    }

    [Fact]
    public void 새_장비는_빈_줄에_이름과_규격을_쓴다()
    {
        var input = DailyReportSample.Input(Thu);
        input.Equipment.Add(new EquipmentEntry { Name = "B/H 6W", Spec = "6t", PrevDayCount = 0, TodayCount = 1 });
        string path = MakeDay(input);

        using var r = new XlsxReader(path);
        r.Value(DailyReportSample.Sheet, "A23").Should().Be("B/H 6W");
        r.Value(DailyReportSample.Sheet, "D23").Should().Be("6t");
        r.Int(DailyReportSample.Sheet, "H23").Should().Be(1);
        r.RawCell(DailyReportSample.Sheet, "J23")!.CellFormula.Should().NotBeNull();

        // 다음 날에는 그 줄을 찾아서 이어 쓴다
        string fri = MakeDay(AddEquipment(DailyReportSample.Input(Fri), "B/H 6W", "6t", 2));
        using var r2 = new XlsxReader(fri);
        r2.Value(DailyReportSample.Sheet, "A23").Should().Be("B/H 6W");
        r2.Int(DailyReportSample.Sheet, "F23").Should().Be(1);
        r2.Int(DailyReportSample.Sheet, "H23").Should().Be(2);
        r2.Value(DailyReportSample.Sheet, "A24").Should().BeNullOrEmpty("같은 장비가 두 줄로 늘면 안 된다");
    }

    private static DailyInput AddEquipment(DailyInput input, string name, string spec, int today)
    {
        input.Equipment.Add(new EquipmentEntry { Name = name, Spec = spec, TodayCount = today });
        return input;
    }

    [Fact]
    public void 빈_줄이_없으면_넣지_못했다고_알린다()
    {
        var input = DailyReportSample.Input(Thu);
        AddEquipment(input, "장비1", "", 1);
        AddEquipment(input, "장비2", "", 1);
        AddEquipment(input, "장비3", "", 1);   // 빈 줄은 23·24 두 개뿐
        var notes = new List<string>();
        MakeDay(input, notes: notes);
        notes.Should().Contain(n => n.Contains("장비3") && n.Contains("빈 줄이 없어"));
    }

    [Fact]
    public void 장비_이름은_앞뒤_공백_대소문자를_무시하고_찾는다()
    {
        var input = DailyReportSample.Input(Thu);
        input.Equipment[0].Name = " b/h 380 ";
        string path = MakeDay(input);
        using var r = new XlsxReader(path);
        r.Int(DailyReportSample.Sheet, "F20").Should().Be(11);
        r.Value(DailyReportSample.Sheet, "A23").Should().BeNullOrEmpty("새 줄로 추가되면 안 된다");
    }

    [Fact]
    public void 같은_이름_장비가_두_줄이면_각각_따로_이어받는다()
    {
        // 예: 덤프 15m3 / 덤프 25t 처럼 이름이 같고 규격만 다른 장비
        var s = new SheetSpec(DailyReportSample.Sheet)
        {
            ["A20"] = "덤프", ["D20"] = "15m3", ["F20"] = 10, ["H20"] = 1,
            ["A21"] = "덤프", ["D21"] = "25t",  ["F21"] = 40, ["H21"] = 4,
        };
        string tpl = XlsxBuilder.Create(_dir.File("dup.xlsx"), s);
        var cfg = DailyReportSample.Config();
        var input = new DailyInput
        {
            Date = Thu,
            Equipment = [new() { Name = "덤프", Spec = "15m3", TodayCount = 2 }, new() { Name = "덤프", Spec = "25t", TodayCount = 3 }],
        };

        GenericExcelWriter.LoadPreviousDay(tpl, input, cfg);

        input.Equipment[0].PrevDayCount.Should().Be(11, "덤프 15m3: 10 + 1");
        input.Equipment[1].PrevDayCount.Should().Be(44, "덤프 25t: 40 + 4");
    }

    [Fact]
    public void 같은_이름_장비는_화면_순서가_달라도_규격으로_제_줄에_쓴다()
    {
        var s = new SheetSpec(DailyReportSample.Sheet)
        {
            ["A20"] = "덤프", ["D20"] = "15m3", ["F20"] = 10, ["H20"] = 1,
            ["A21"] = "덤프", ["D21"] = "25t",  ["F21"] = 40, ["H21"] = 4,
        };
        XlsxBuilder.Create(_dir.File(DailyReportSample.TemplateName), s);
        var input = new DailyInput
        {
            Date = Thu,
            // 화면에서는 25t 가 먼저
            Equipment = [new() { Name = "덤프", Spec = "25t", TodayCount = 3 }, new() { Name = "덤프", Spec = "15m3", TodayCount = 2 }],
        };

        string path = MakeDay(input);

        using var r = new XlsxReader(path);
        r.Int(DailyReportSample.Sheet, "F20").Should().Be(11);
        r.Int(DailyReportSample.Sheet, "H20").Should().Be(2, "15m3 줄");
        r.Int(DailyReportSample.Sheet, "F21").Should().Be(44);
        r.Int(DailyReportSample.Sheet, "H21").Should().Be(3, "25t 줄");
    }

    // ── 인원 표 ──────────────────────────────────────────────────────

    [Fact]
    public void 인원_전일누계가_이틀_연속_이어진다()
    {
        MakeDay(DailyReportSample.Input(Thu, manager: 2, scaffold: 6));
        string fri = MakeDay(DailyReportSample.Input(Fri, manager: 3, scaffold: 4));
        using var r = new XlsxReader(fri);
        r.Int(DailyReportSample.Sheet, "N20").Should().Be(35, "관리자 30+3 → 33, +2");
        r.Int(DailyReportSample.Sheet, "N21").Should().Be(61, "비계 50+5 → 55, +6");
        r.Int(DailyReportSample.Sheet, "P21").Should().Be(4);
    }

    [Fact]
    public void 수장재를_포함하는_문서는_수장재_줄도_쓴다()
    {
        string path = MakeDay(DailyReportSample.Input(Thu, sujang: 4), DailyReportSample.Config(includesSujangjae: true));
        using var r = new XlsxReader(path);
        r.Int(DailyReportSample.Sheet, "N22").Should().Be(23, "20 + 3");
        r.Int(DailyReportSample.Sheet, "P22").Should().Be(4);
    }

    [Fact]
    public void 수장재를_빼는_문서는_수장재_줄_금일이_0이고_누계는_이월()
    {
        // 지금 규칙: 공사일보(IncludesSujangjae=false)는 수장재를 화면 목록에 넣지 않으므로 "지운 줄"처럼 처리한다.
        // (수장재가 공사일보에 들어가는지는 아직 미결 — docs/기획.md)
        string path = MakeDay(DailyReportSample.Input(Thu, sujang: 4), DailyReportSample.Config(includesSujangjae: false));
        using var r = new XlsxReader(path);
        r.Int(DailyReportSample.Sheet, "N22").Should().Be(23);
        r.Int(DailyReportSample.Sheet, "P22").Should().Be(0);
    }

    [Fact]
    public void 수장재_전일누계도_앞_파일에서_읽는다()
    {
        var input = DailyReportSample.Input(Thu);
        GenericExcelWriter.LoadPreviousDay(_template, input, DailyReportSample.Config());
        input.SujangjaePersonnel[0].PrevCumulativeCount.Should().Be(23);
    }

    // ── 표 줄 목록 읽기 ─────────────────────────────────────────────

    [Fact]
    public void 파일의_장비_직종_줄을_읽는다()
    {
        var (eq, roles) = GenericExcelWriter.ReadTableRows(_template, DailyReportSample.Config());
        eq.Should().Equal(("B/H 380", "39.6t"), ("덤프", "15m3"), ("세륜기", ""));
        roles.Should().Equal("관리자", "비계", "수장재");
    }

    // ── 바탕 파일·백업 ──────────────────────────────────────────────

    [Fact]
    public void 앞_날짜_파일을_바탕으로_만들어_지정_안_한_칸이_이어진다()
    {
        string thu = MakeDay(DailyReportSample.Input(Thu));
        // 사용자가 목요일 파일의 지정 안 한 칸(특기사항 등)을 손으로 고쳤다고 치자
        using (var d = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(thu, true))
        {
            var ws = CellHelper.FindWorksheetPart(d.WorkbookPart!, DailyReportSample.Sheet)!;
            CellHelper.SetCellString(d, ws, "B40", "특기사항: 우천 대비");
        }

        var notes = new List<string>();
        string fri = MakeDay(DailyReportSample.Input(Fri), notes: notes);

        using var r = new XlsxReader(fri);
        r.Value(DailyReportSample.Sheet, "B40").Should().Be("특기사항: 우천 대비");
        notes.Should().Contain(n => n.Contains("11/05 파일을 바탕으로"));
    }

    [Fact]
    public void 앞_날짜_바탕을_끄면_항상_서식에서_만든다()
    {
        var cfg = DailyReportSample.Config(usePrevious: false);
        string thu = MakeDay(DailyReportSample.Input(Thu), cfg);
        using (var d = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(thu, true))
            CellHelper.SetCellString(d, CellHelper.FindWorksheetPart(d.WorkbookPart!, DailyReportSample.Sheet)!, "B40", "손으로 쓴 값");

        string fri = MakeDay(DailyReportSample.Input(Fri), cfg);
        using var r = new XlsxReader(fri);
        r.Value(DailyReportSample.Sheet, "B40").Should().BeNull();
    }

    [Fact]
    public void 같은_날_다시_만들면_앞_파일을_백업하고_어제_파일을_바탕으로_한다()
    {
        MakeDay(DailyReportSample.Input(Thu, bh: 2));
        MakeDay(DailyReportSample.Input(Fri, bh: 4));
        string fri = MakeDay(DailyReportSample.Input(Fri, bh: 6));   // 금요일 다시 생성

        Directory.GetFiles(Path.Combine(_dir.Path, Out, OutputBackup.FolderName)).Should().ContainSingle();
        using var r = new XlsxReader(fri);
        r.Int(DailyReportSample.Sheet, "F20").Should().Be(13, "금요일 자신이 아니라 목요일 파일에서 이어받아야 한다");
        r.Int(DailyReportSample.Sheet, "H20").Should().Be(6);
    }

    [Fact]
    public void 서식이_없으면_알기_쉬운_오류()
    {
        var cfg = DailyReportSample.Config();
        cfg.TemplateFile = "없는파일.xlsx";
        var act = () => GenericExcelWriter.Generate(DailyReportSample.Input(Thu), cfg, _dir.Path, Out);
        act.Should().Throw<FileNotFoundException>().WithMessage("*템플릿 파일*");
    }

    [Fact]
    public void 시트_이름이_틀리면_알기_쉬운_오류()
    {
        var cfg = DailyReportSample.Config();
        cfg.SheetName = "TBM일지 서식12월";
        var act = () => GenericExcelWriter.Generate(DailyReportSample.Input(Thu), cfg, _dir.Path, Out);
        act.Should().Throw<InvalidOperationException>().WithMessage("*TBM일지 서식12월*");
    }

    // ── 합계 대조 (수식과 별개로 직접 계산) ─────────────────────────

    [Fact]
    public void 금일_합계를_직접_계산하면_입력과_같다()
    {
        var input = DailyReportSample.Input(Thu, bh: 2, dump: 3, wash: 1, manager: 3, scaffold: 7);
        string path = MakeDay(input);
        using var r = new XlsxReader(path);

        int eqToday = Enumerable.Range(20, 5).Sum(row => r.Int(DailyReportSample.Sheet, $"H{row}") ?? 0);
        eqToday.Should().Be(input.Equipment.Sum(e => e.TodayCount));

        int eqCum = Enumerable.Range(20, 5).Sum(row =>
            (r.Int(DailyReportSample.Sheet, $"F{row}") ?? 0) + (r.Int(DailyReportSample.Sheet, $"H{row}") ?? 0));
        eqCum.Should().Be(11 + 7 + 4 + 6, "서식 누계 22 + 오늘 6");

        int pnToday = Enumerable.Range(20, 5).Sum(row => r.Int(DailyReportSample.Sheet, $"P{row}") ?? 0);
        pnToday.Should().Be(input.TotalPersonnelWithoutSujangjae);
    }
}
