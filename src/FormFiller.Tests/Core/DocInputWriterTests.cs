using FluentAssertions;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Core.Models;
using FormFiller.Tests.Helpers;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace FormFiller.Tests.Core;

/// <summary>
/// TBM 일지 칸, 매일 직접 입력(측정값), 사진 교체, 그림 위치 옮기기 테스트.
/// </summary>
public class DocInputWriterTests : IDisposable
{
    private const string Out = "output";
    private static readonly DateTime Thu = new(2026, 11, 5);
    private readonly TempDir _dir = new();

    public void Dispose() => _dir.Dispose();

    // ── TBM 일지 ─────────────────────────────────────────────────────

    private DocumentConfig Tbm() => new()
    {
        Id = "tbm", Name = "TBM", TemplateFile = "tbm.xlsx", OutputPattern = "{yy}.{MM}.{dd} TBM.xlsx",
        SheetName = "TBM일지 서식11월",
        Fields =
        [
            new() { Id = "date",       Type = FieldTypes.DateSimple,           Cell = "F4" },
            new() { Id = "personnel",  Type = FieldTypes.TbmPersonnel,         Cell = "AD6" },
            new() { Id = "equipment",  Type = FieldTypes.ActiveEquipmentNames, Cell = "F7" },
            new() { Id = "planned",    Type = FieldTypes.TbmPersonnel,         Cell = "I16" },
            new() { Id = "actual",     Type = FieldTypes.TbmPersonnel,         Cell = "O16" },
            new() { Id = "rate",       Type = FieldTypes.AttendanceRate,       Cell = "U16" },
        ],
    };

    [Fact]
    public void TBM_작성일자_인원_장비_참석율()
    {
        XlsxBuilder.Create(_dir.File("tbm.xlsx"), new SheetSpec("TBM일지 서식11월")
        {
            ["F4"] = "2026.11.04", ["AD6"] = 14, ["F7"] = "B/H 380", ["I16"] = 14, ["O16"] = 14, ["U16"] = 100,
        });
        var input = DailyReportSample.Input(Thu, bh: 1, dump: 0, wash: 2, manager: 3, scaffold: 8, sujang: 4);

        string path = GenericExcelWriter.Generate(input, Tbm(), _dir.Path, Out);

        using var r = new XlsxReader(path);
        const string s = "TBM일지 서식11월";
        r.Value(s, "F4").Should().Be("2026.11.05", "TBM 작성일자는 점 구분 문자열");
        r.Int(s, "AD6").Should().Be(15, "공사 3+8 + 수장재 4");
        r.Int(s, "I16").Should().Be(15);
        r.Int(s, "O16").Should().Be(15);
        r.Int(s, "U16").Should().Be(100);
        r.Value(s, "F7").Should().Be("B/H 380, 세륜기", "금일 0대인 덤프는 빠진다");
        r.ValidationErrors().Should().BeEmpty();
    }

    // ── 매일 직접 입력 (시트마다 따로) ─────────────────────────────

    private DocumentConfig Noise() => new()
    {
        Id = "noise", Name = "소음측정일지", TemplateFile = "noise.xlsx", OutputPattern = "{yy}.{MM}.{dd} 소음.xlsx",
        SheetNames = ["오전초등학교", "다사랑병원"],
        Fields =
        [
            new() { Id = "date", Type = FieldTypes.DateKorean, Cell = "B2" },
            new() { Id = "level", Label = "측정값", Type = FieldTypes.Input, Cell = "E5", Suffix = "dB" },
            new() { Id = "photo1", Label = "측정사진", Type = FieldTypes.Photo, Cell = "D12", ValueCell = "G12", Suffix = "dB" },
            new()
            {
                Id = "pos", Label = "장비 위치", Type = FieldTypes.ShapePosition,
                Presets =
                [
                    new ShapePreset
                    {
                        Name = "정문 앞",
                        Shapes = [new ShapePlacement { Name = "TextBox 9", Index = 1, FromCol = 10, FromRow = 30, ToCol = 12, ToRow = 32 }],
                    },
                ],
            },
        ],
    };

    private void CreateNoiseTemplate()
    {
        SheetSpec Sheet(string name, string picName) => new(name)
        {
            ["B2"] = "옛 날짜", ["E5"] = "50.1dB", ["G12"] = "48.0dB",
            Pictures = { new Pic("D12", picName) },
            Boxes = { new Box("H20", "TextBox 9") },
        };
        // 같은 모양 시트끼리 그림 이름이 다르다 (실제 서식에서 확인한 모양)
        XlsxBuilder.Create(_dir.File("noise.xlsx"), Sheet("오전초등학교", "그림 21"), Sheet("다사랑병원", "그림 22"));
    }

    [Fact]
    public void 직접_입력값은_시트마다_따로_쓰고_단위를_붙인다()
    {
        CreateNoiseTemplate();
        var input = new DailyInput { Date = Thu };
        input.DocValues[DailyInput.Key("noise", "오전초등학교", "level")] = "55.3";
        input.DocValues[DailyInput.Key("noise", "다사랑병원", "level")] = "61.0dB";   // 이미 붙어 있으면 그대로

        string path = GenericExcelWriter.Generate(input, Noise(), _dir.Path, Out);

        using var r = new XlsxReader(path);
        r.Value("오전초등학교", "E5").Should().Be("55.3dB");
        r.Value("다사랑병원", "E5").Should().Be("61.0dB");
        r.Value("다사랑병원", "B2").Should().Be("2026년 11월 05일 목요일");
    }

    [Fact]
    public void 입력이_없으면_바탕_파일_값을_두고_알린다()
    {
        CreateNoiseTemplate();
        var notes = new List<string>();
        string path = GenericExcelWriter.Generate(new DailyInput { Date = Thu }, Noise(), _dir.Path, Out, notes: notes);

        using var r = new XlsxReader(path);
        r.Value("오전초등학교", "E5").Should().Be("50.1dB");
        notes.Should().Contain(n => n.Contains("[오전초등학교] 측정값: 입력 없음"));
        notes.Should().Contain(n => n.Contains("측정사진: 사진 없음"));
    }

    [Theory]
    [InlineData("55.3", "dB", "55.3dB")]
    [InlineData(" 55.3 ", "dB", "55.3dB")]
    [InlineData("55.3DB", "dB", "55.3DB")]
    [InlineData("", "dB", "")]
    [InlineData("12", null, "12")]
    public void 단위_붙이기(string value, string? suffix, string expected) =>
        GenericExcelWriter.WithSuffix(value, suffix).Should().Be(expected);

    // ── 사진 교체 ────────────────────────────────────────────────────

    private (string Path, List<(string Photo, int W, int H)> Calls) GenerateWithPhoto(string sheetWithPhoto)
    {
        CreateNoiseTemplate();
        string photo = _dir.File("p.jpg");
        File.WriteAllBytes(photo, XlsxBuilder.FakeJpeg);
        var input = new DailyInput { Date = Thu };
        input.DocValues[DailyInput.Key("noise", sheetWithPhoto, "photo1")] = photo;
        input.DocValues[DailyInput.Key("noise", sheetWithPhoto, "photo1" + DailyInput.ValueSuffixKey)] = "57";

        var calls = new List<(string, int, int)>();
        string path = GenericExcelWriter.Generate(input, Noise(), _dir.Path, Out,
            photoLoader: (p, w, h) => { calls.Add((p, w, h)); return XlsxBuilder.FakeJpeg; });
        return (path, calls);
    }

    [Fact]
    public void 사진은_그_칸의_그림만_바꾸고_위치와_크기는_그대로()
    {
        var (path, calls) = GenerateWithPhoto("다사랑병원");

        // 그림 크기 1,905,000 × 1,428,750 EMU = 200 × 150 px → 인쇄용 2배
        calls.Should().ContainSingle().Which.Should().Be((_dir.File("p.jpg"), 400, 300));

        using var r = new XlsxReader(path);
        var anchors = r.Anchors("다사랑병원");
        anchors.Should().HaveCount(2, "그림·글상자 개수는 그대로");
        var pic = anchors.Single(a => a.Descendants<Xdr.Picture>().Any());
        pic.FromMarker!.ColumnId!.Text.Should().Be("3");
        pic.FromMarker.RowId!.Text.Should().Be("11");
        pic.Descendants<A.Extents>().First().Cx!.Value.Should().Be(1_905_000);

        string rid = pic.Descendants<A.Blip>().Single().Embed!.Value!;
        var img = r.Sheet("다사랑병원").DrawingsPart!.GetPartById(rid);
        using (var st = img.GetStream()) new BinaryReader(st).ReadBytes(13).Should().Equal(XlsxBuilder.FakeJpeg);

        r.Value("다사랑병원", "G12").Should().Be("57dB", "사진 옆 측정값");
        r.Value("오전초등학교", "G12").Should().Be("48.0dB", "다른 시트는 그대로");
        r.ValidationErrors().Should().BeEmpty();
    }

    [Fact]
    public void 사진을_바꾸면_안_쓰는_옛_이미지는_지운다()
    {
        var (path, _) = GenerateWithPhoto("오전초등학교");
        using var r = new XlsxReader(path);
        r.Sheet("오전초등학교").DrawingsPart!.ImageParts.Should().ContainSingle("옛 PNG 는 지우고 새 JPEG 하나만");
        r.Sheet("다사랑병원").DrawingsPart!.ImageParts.Should().ContainSingle();
    }

    [Fact]
    public void 사진_칸이_한_줄_어긋나도_같은_열이면_찾는다()
    {
        // 손으로 붙인 사진이 D12 가 아니라 D11 에서 시작하는 경우
        XlsxBuilder.Create(_dir.File("noise.xlsx"),
            new SheetSpec("오전초등학교") { Pictures = { new Pic("D11", "그림 1") } },
            new SheetSpec("다사랑병원"));
        using var r = new XlsxReader(_dir.File("noise.xlsx"));
        ExcelDrawingHelper.FindPictureAt(r.Sheet("오전초등학교"), "D12").Should().NotBeNull();
        ExcelDrawingHelper.FindPictureAt(r.Sheet("오전초등학교"), "D14").Should().BeNull("두 줄 차이는 다른 칸");
        ExcelDrawingHelper.FindPictureAt(r.Sheet("오전초등학교"), "E12").Should().BeNull("다른 열");
    }

    [Fact]
    public void 그림_없는_칸에_사진을_넣으려면_알린다()
    {
        XlsxBuilder.Create(_dir.File("noise.xlsx"), new SheetSpec("오전초등학교"), new SheetSpec("다사랑병원"));
        string photo = _dir.File("p.jpg");
        File.WriteAllBytes(photo, XlsxBuilder.FakeJpeg);
        var input = new DailyInput { Date = Thu };
        input.DocValues[DailyInput.Key("noise", "오전초등학교", "photo1")] = photo;
        var notes = new List<string>();

        GenericExcelWriter.Generate(input, Noise(), _dir.Path, Out, photoLoader: (_, _, _) => XlsxBuilder.FakeJpeg, notes: notes);

        notes.Should().Contain(n => n.Contains("D12 칸에 그림이 없어"));
    }

    // ── 그림 위치 ────────────────────────────────────────────────────

    [Fact]
    public void 저장한_위치를_고르면_그림이_그_자리로_옮겨진다()
    {
        CreateNoiseTemplate();
        var input = new DailyInput { Date = Thu };
        input.DocValues[DailyInput.Key("noise", "다사랑병원", "pos")] = "정문 앞";

        string path = GenericExcelWriter.Generate(input, Noise(), _dir.Path, Out);

        using var r = new XlsxReader(path);
        var moved = ExcelDrawingHelper.ReadPlacement(r.Sheet("다사랑병원"), "TextBox 9")!;
        (moved.FromCol, moved.FromRow, moved.ToCol, moved.ToRow).Should().Be((10, 30, 12, 32));
        var same = ExcelDrawingHelper.ReadPlacement(r.Sheet("오전초등학교"), "TextBox 9")!;
        (same.FromCol, same.FromRow).Should().Be((7, 19), "고르지 않은 시트는 그대로");
    }

    [Fact]
    public void 없는_위치_이름이면_알린다()
    {
        CreateNoiseTemplate();
        var input = new DailyInput { Date = Thu };
        input.DocValues[DailyInput.Key("noise", "오전초등학교", "pos")] = "후문";
        var notes = new List<string>();
        GenericExcelWriter.Generate(input, Noise(), _dir.Path, Out, notes: notes);
        notes.Should().Contain(n => n.Contains("저장된 위치 '후문' 없음"));
    }

    [Theory]
    [InlineData("그림 21", "Picture 21", true)]    // 한국어 파일 이름 ↔ Excel 화면 이름
    [InlineData("그림 21", "그림 21", true)]
    [InlineData("그림 21", "Picture 2", false)]
    [InlineData("TextBox 10", "텍스트 상자 10", true)]
    [InlineData(null, "그림 1", false)]
    public void 그림_이름_비교(string? fileName, string wanted, bool expected) =>
        ExcelDrawingHelper.NameMatches(fileName, wanted).Should().Be(expected);
}
