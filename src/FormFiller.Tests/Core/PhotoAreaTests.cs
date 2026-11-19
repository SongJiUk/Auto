using FluentAssertions;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Core.Models;
using FormFiller.Tests.Helpers;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace FormFiller.Tests.Core;

/// <summary>
/// 사진 구역 블록 (BlockTypes.Photos).
/// TBM 사진대지처럼 "사진 칸 2개 + 칸마다 설명 줄" 서식에서, 그날 사진 장수(0~2)에 맞춰
/// 구역 안 그림을 지우고 정한 자리에 새로 넣는다. 1장일 때 자리·설명 칸은 문서마다 설정한다.
/// </summary>
public class PhotoAreaTests : IDisposable
{
    private const string Out = "output", Sheet = "사진대지";
    private static readonly DateTime Mon = new(2026, 11, 16), Tue = new(2026, 11, 17);
    private readonly TempDir _dir = new();
    private readonly List<(string Path, int W, int H)> _loaded = [];
    public void Dispose() => _dir.Dispose();

    // 사진 칸 2개 (위·아래) + 1장일 때 두 칸을 합친 자리
    private static readonly PhotoArea Top    = new() { Range = "A4:H15", CaptionCell = "B16" };
    private static readonly PhotoArea Bottom = new() { Range = "A17:H28", CaptionCell = "B29" };
    private static readonly PhotoArea Whole  = new() { Range = "A4:H28", CaptionCell = "B29" };

    private static FieldConfig PhotoBlock(PhotoArea? single = null) => new()
    {
        Id = "TBM사진", Label = "TBM사진", Type = BlockTypes.Photos,
        PhotoAreas = [Top, Bottom], SinglePhotoArea = single,
    };

    private DocumentConfig Doc(string template, FieldConfig block) => new()
    {
        Id = "tbm", Name = "TBM", TemplateFile = template,
        OutputPattern = "{yy}.{MM}.{dd} TBM.xlsx", SheetName = Sheet, Fields = [block],
    };

    /// <summary>서식: 설명 줄 2개, 날짜 수식, 구역 밖 도장 그림(K2) + 구역 안 그림들</summary>
    private string Template(params Pic[] areaPictures)
    {
        var s = new SheetSpec(Sheet)
        {
            ["B16"] = "옛 설명1", ["B29"] = "옛 설명2",
            ["F16"] = new F("TODAY()", "46000"), ["F29"] = new F("TODAY()", "46000"),
        };
        s.Pictures.Add(new Pic("K2", "도장"));
        s.Pictures.AddRange(areaPictures);
        return XlsxBuilder.Create(_dir.File("tbm.xlsx"), s);
    }

    private string Photo(string name)
    {
        string p = _dir.File(name);
        File.WriteAllBytes(p, XlsxBuilder.FakeJpeg);
        return p;
    }

    private (string Path, List<string> Notes, List<string> Problems) Generate(DocumentConfig doc, DateTime date,
                                                                             params (string Path, string Caption)[] photos)
    {
        var input = new DailyInput { Date = date };
        input.Blocks["TBM사진"] = new BlockValue
        {
            Photos = [.. photos.Select(p => new PhotoValue { Path = p.Path, Caption = p.Caption })],
        };
        var notes = new List<string>();
        var problems = new List<string>();
        string path = GenericExcelWriter.Generate(input, doc, _dir.Path, Out,
            photoLoader: (p, w, h) => { _loaded.Add((p, w, h)); return XlsxBuilder.FakeJpeg; },
            notes: notes, problems: problems);
        return (path, notes, problems);
    }

    /// <summary>구역 밖(도장)을 뺀 그림들의 (from 열, from 행, to 열, to 행) — 0부터 세는 번호</summary>
    private static List<(int FC, int FR, int TC, int TR)> AreaPictures(XlsxReader r) =>
        r.Anchors(Sheet).Where(a => a.Descendants<Xdr.Picture>().Any())
         .Select(a => (FC: int.Parse(a.FromMarker!.ColumnId!.Text), FR: int.Parse(a.FromMarker.RowId!.Text),
                       TC: int.Parse(a.ToMarker!.ColumnId!.Text),   TR: int.Parse(a.ToMarker.RowId!.Text)))
         .Where(p => p != (10, 1, 12, 6))          // 도장 (K2)
         .OrderBy(p => p.FR).ToList();

    [Fact]
    public void Generate_사진2장_칸마다그림과설명()
    {
        // 서식은 1장 모양 (두 칸을 덮는 큰 그림 하나)
        var doc = Doc(Template(new Pic("A4", "그림 2")), PhotoBlock(Whole));

        var (path, _, problems) = Generate(doc, Mon, (Photo("1.jpg"), "일일 TBM(수장재)"), (Photo("2.jpg"), "일일 TBM(장비)"));

        problems.Should().BeEmpty("사진 구역 안 그림 수가 바뀌는 건 정상");
        using var r = new XlsxReader(path);
        AreaPictures(r).Should().Equal((0, 3, 8, 15), (0, 16, 8, 28));
        r.Value(Sheet, "B16").Should().Be("일일 TBM(수장재)");
        r.Value(Sheet, "B29").Should().Be("일일 TBM(장비)");
        r.RawCell(Sheet, "F16")!.CellFormula.Should().NotBeNull("날짜 수식은 그대로");
        r.ValidationErrors().Should().BeEmpty();
        _loaded.Select(l => Path.GetFileName(l.Path)).Should().Equal("1.jpg", "2.jpg");
    }

    [Fact]
    public void Generate_사진1장_한장자리에꽉채우고_정한칸에설명()
    {
        // 서식은 2장 모양 (칸마다 그림)
        var doc = Doc(Template(new Pic("A4", "그림 3"), new Pic("A17", "그림 2")), PhotoBlock(Whole));

        var (path, _, problems) = Generate(doc, Mon, (Photo("1.jpg"), "다 같이 TBM"));

        problems.Should().BeEmpty();
        using var r = new XlsxReader(path);
        AreaPictures(r).Should().Equal((0, 3, 8, 28));
        r.Value(Sheet, "B29").Should().Be("다 같이 TBM", "1장일 때 설명 칸 (설정)");
        r.Value(Sheet, "B16").Should().Be("옛 설명1", "다른 설명 칸은 건드리지 않음 (그림에 가려짐)");
    }

    [Fact]
    public void Generate_사진1장_한장자리설정없으면_첫칸()
    {
        var doc = Doc(Template(new Pic("A4", "그림 3"), new Pic("A17", "그림 2")), PhotoBlock(single: null));

        var (path, _, _) = Generate(doc, Mon, (Photo("1.jpg"), "수장재만"));

        using var r = new XlsxReader(path);
        AreaPictures(r).Should().Equal((0, 3, 8, 15));
        r.Value(Sheet, "B16").Should().Be("수장재만");
    }

    [Fact]
    public void Generate_한장자리_설명칸이비면_설명은안씀()
    {
        var doc = Doc(Template(new Pic("A4", "그림 2")), PhotoBlock(new PhotoArea { Range = "A4:H28" }));

        var (path, _, _) = Generate(doc, Mon, (Photo("1.jpg"), "설명"));

        using var r = new XlsxReader(path);
        r.Value(Sheet, "B16").Should().Be("옛 설명1");
        r.Value(Sheet, "B29").Should().Be("옛 설명2");
    }

    [Fact]
    public void Generate_사진없으면_바탕파일그대로_안내()
    {
        var doc = Doc(Template(new Pic("A4", "그림 2")), PhotoBlock(Whole));

        var (path, notes, problems) = Generate(doc, Mon);

        problems.Should().BeEmpty();
        using var r = new XlsxReader(path);
        AreaPictures(r).Should().Equal(new[] { (0, 3, 2, 8) }, "서식 그림 그대로 (XlsxBuilder 기본 크기)");
        notes.Should().Contain(n => n.Contains("TBM사진") && n.Contains("사진 없음"));
    }

    [Fact]
    public void Generate_칸보다사진이많으면_칸수만큼_안내()
    {
        var doc = Doc(Template(new Pic("A4", "그림 2")), PhotoBlock(Whole));

        var (path, notes, _) = Generate(doc, Mon,
            (Photo("1.jpg"), "a"), (Photo("2.jpg"), "b"), (Photo("3.jpg"), "c"));

        using var r = new XlsxReader(path);
        AreaPictures(r).Should().HaveCount(2);
        notes.Should().Contain(n => n.Contains("3장") && n.Contains("2칸"));
    }

    [Fact]
    public void Generate_없는사진파일은_건너뜀()
    {
        var doc = Doc(Template(new Pic("A4", "그림 2")), PhotoBlock(Whole));

        var (path, notes, _) = Generate(doc, Mon, (_dir.File("없음.jpg"), "a"), (Photo("2.jpg"), "b"));

        using var r = new XlsxReader(path);
        AreaPictures(r).Should().Equal(new[] { (0, 3, 8, 28) }, "남은 사진 1장 → 1장 자리");
        notes.Should().Contain(n => n.Contains("없음.jpg"));
    }

    [Fact]
    public void Generate_구역밖그림은_그대로()
    {
        var doc = Doc(Template(new Pic("A4", "그림 2")), PhotoBlock(Whole));

        var (path, _, _) = Generate(doc, Mon, (Photo("1.jpg"), "a"), (Photo("2.jpg"), "b"));

        using var r = new XlsxReader(path);
        r.Anchors(Sheet).Count(a => a.Descendants<Xdr.Picture>().Any()).Should().Be(3, "도장 + 사진 2장");
        ExcelDrawingHelper.FindPictureAt(r.Sheet(Sheet), "K2").Should().NotBeNull();
    }

    [Fact]
    public void Generate_사진크기_칸모양대로_가로세로비율()
    {
        var doc = Doc(Template(new Pic("A4", "그림 2")), PhotoBlock(Whole));

        Generate(doc, Mon, (Photo("1.jpg"), "a"));
        Generate(doc, Tue, (Photo("1.jpg"), "a"), (Photo("2.jpg"), "b"));

        // 기본 열 너비 8.43 → 64px, 기본 행 높이 15pt → 20px, 인쇄용 2배 해상도
        (_loaded[0].W, _loaded[0].H).Should().Be((8 * 64 * 2, 25 * 20 * 2), "1장 자리 A4:H28");
        (_loaded[1].W, _loaded[1].H).Should().Be((8 * 64 * 2, 12 * 20 * 2), "위 칸 A4:H15");
        (_loaded[2].W, _loaded[2].H).Should().Be((8 * 64 * 2, 12 * 20 * 2), "아래 칸 A17:H28");
    }

    [Fact]
    public void Generate_사진크기_열너비행높이를바꾼서식()
    {
        string template = Template(new Pic("A4", "그림 2"));
        using (var pkg = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(template, true))
        {
            var ws = CellHelper.FindWorksheetPart(pkg.WorkbookPart!, Sheet)!.Worksheet;
            var data = ws.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()!;
            ws.InsertBefore(new DocumentFormat.OpenXml.Spreadsheet.Columns(
                new DocumentFormat.OpenXml.Spreadsheet.Column { Min = 1, Max = 1, Width = 20, CustomWidth = true }), data);
            data.InsertAt(new DocumentFormat.OpenXml.Spreadsheet.Row { RowIndex = 4, Height = 30, CustomHeight = true }, 0);
            ws.Save();
        }
        var doc = Doc("tbm.xlsx", PhotoBlock(Whole));

        Generate(doc, Mon, (Photo("1.jpg"), "a"), (Photo("2.jpg"), "b"));

        // A열 너비 20 → Truncate((256×20 + 18) / 256 × 7) = 140px + B~H 7×64px, 4행 30pt → 40px + 5~15행 11×20px
        (_loaded[0].W, _loaded[0].H).Should().Be(((140 + 7 * 64) * 2, (40 + 11 * 20) * 2));
    }

    [Fact]
    public void Generate_사진크기_맑은고딕서식()
    {
        // 실제 TBM 서식과 같은 조건: 기본 글꼴 맑은 고딕 11, 기본 열 너비 10.5, 기본 행 높이 25.5pt
        string template = Template(new Pic("A4", "그림 2"));
        using (var pkg = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(template, true))
        {
            var styles = pkg.WorkbookPart!.AddNewPart<DocumentFormat.OpenXml.Packaging.WorkbookStylesPart>();
            styles.Stylesheet = new DocumentFormat.OpenXml.Spreadsheet.Stylesheet(
                new DocumentFormat.OpenXml.Spreadsheet.Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font(
                    new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 11 },
                    new DocumentFormat.OpenXml.Spreadsheet.FontName { Val = "맑은 고딕" })) { Count = 1 });
            var ws = CellHelper.FindWorksheetPart(pkg.WorkbookPart!, Sheet)!.Worksheet;
            ws.InsertAt(new DocumentFormat.OpenXml.Spreadsheet.SheetFormatProperties
                { DefaultColumnWidth = 10.5, DefaultRowHeight = 25.5, CustomHeight = true }, 0);
            ws.Save();
        }
        var doc = Doc("tbm.xlsx", PhotoBlock(Whole));

        Generate(doc, Mon, (Photo("1.jpg"), "a"), (Photo("2.jpg"), "b"));

        // 열 = Truncate((256×10.5 + 16) / 256 × 8) = 84px, 행 = 25.5pt × 4/3 = 34px
        // → 칸 하나 672×408px = 17.8×10.8cm (엑셀이 저장한 실제 그림 17.68×10.83cm)
        (_loaded[0].W, _loaded[0].H).Should().Be((8 * 84 * 2, 12 * 34 * 2));
    }

    [Fact]
    public void Generate_이틀연속_옛이미지를지워_파일이커지지않음()
    {
        var doc = Doc(Template(new Pic("A4", "그림 2")), PhotoBlock(Whole));

        var (mon, _, _) = Generate(doc, Mon, (Photo("1.jpg"), "a"), (Photo("2.jpg"), "b"));
        var (tue, _, problems) = Generate(doc, Tue, (Photo("3.jpg"), "c"));

        problems.Should().BeEmpty("앞 날짜(2장) 파일을 바탕으로 1장을 넣어도 정상");
        using var r = new XlsxReader(tue);
        r.Sheet(Sheet).DrawingsPart!.ImageParts.Count().Should().Be(2, "도장 + 오늘 사진 1장");
        AreaPictures(r).Should().Equal((0, 3, 8, 28));
        r.ValidationErrors().Should().BeEmpty();
    }

    [Fact]
    public void Generate_그림이하나도없는시트에도_넣음()
    {
        var s = new SheetSpec(Sheet) { ["B16"] = "옛 설명1", ["B29"] = "옛 설명2" };
        var doc = Doc(XlsxBuilder.Create(_dir.File("tbm.xlsx"), s), PhotoBlock(Whole));

        var (path, _, _) = Generate(doc, Mon, (Photo("1.jpg"), "a"), (Photo("2.jpg"), "b"));

        using var r = new XlsxReader(path);
        AreaPictures(r).Should().Equal((0, 3, 8, 15), (0, 16, 8, 28));
        r.ValidationErrors().Should().BeEmpty();
    }

    [Fact]
    public void ReadValues_어제설명을_처음값으로()
    {
        var doc = Doc(Template(new Pic("A4", "그림 2")), PhotoBlock(Whole));
        var (mon, _, _) = Generate(doc, Mon, (Photo("1.jpg"), "수장재"), (Photo("2.jpg"), "장비"));

        var init = BlockEngine.InitialValues(doc, BlockEngine.ReadValues(mon, doc));

        init["TBM사진"].Photos.Select(p => p.Caption).Should().Equal("수장재", "장비");
        init["TBM사진"].Photos.Should().OnlyContain(p => p.Path == "", "사진은 매일 새로 고른다");
    }
}
