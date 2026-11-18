using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Core.Models;
using FormFiller.Tests.Helpers;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace FormFiller.Tests.Core;

/// <summary>저장 후 자동 검증 테스트: 정상 파일은 통과, 일부러 망가뜨린 파일은 잡아낸다.</summary>
public class OutputVerifierTests : IDisposable
{
    private readonly TempDir _dir = new();
    public void Dispose() => _dir.Dispose();

    private static readonly DocumentConfig Doc = new()
    {
        Id = "n", Name = "n", SheetName = "S",
        Fields =
        [
            new()
            {
                Id = "pos", Type = FieldTypes.ShapePosition,
                Presets = [new ShapePreset { Name = "p", Shapes = [new ShapePlacement { Index = 1 }] }],
            },
        ],
    };

    /// <summary>그림 1개 + 글상자 1개 시트. 서식 역할</summary>
    private string Base()
    {
        string path = _dir.File("base.xlsx");
        XlsxBuilder.Create(path, new SheetSpec("S")
        {
            ["A1"] = "x",
            Pictures = { new Pic("D12", "그림 1") },
            Boxes = { new Box("H20", "TextBox 2") },
        });
        return path;
    }

    private string CopyAndEdit(string src, Action<SpreadsheetDocument> edit)
    {
        string path = _dir.File($"out_{Guid.NewGuid():N}.xlsx");
        File.Copy(src, path);
        using (var d = SpreadsheetDocument.Open(path, true)) edit(d);
        return path;
    }

    private static WorksheetPart Ws(SpreadsheetDocument d) => CellHelper.FindWorksheetPart(d.WorkbookPart!, "S")!;

    [Fact]
    public void 엔진으로_만든_정상_파일은_통과()
    {
        DailyReportSample.CreateTemplate(_dir.File(DailyReportSample.TemplateName));
        var problems = new List<string>();
        GenericExcelWriter.Generate(DailyReportSample.Input(new DateTime(2026, 11, 5)), DailyReportSample.Config(),
                                    _dir.Path, "output", problems: problems);
        problems.Should().BeEmpty();
    }

    [Fact]
    public void 값만_바꾼_파일은_통과()
    {
        string b = Base();
        string o = CopyAndEdit(b, d => CellHelper.SetCellString(d, Ws(d), "A1", "새 값"));
        OutputVerifier.Verify(b, o, Doc).Should().BeEmpty();
    }

    [Fact]
    public void 그림이_사라지면_잡는다()
    {
        string b = Base();
        string o = CopyAndEdit(b, d => Ws(d).DrawingsPart!.WorksheetDrawing.Elements<Xdr.TwoCellAnchor>().First().Remove());
        OutputVerifier.Verify(b, o, Doc).Should().Contain(p => p.Contains("개수가 달라졌습니다 (2개 → 1개)"));
    }

    [Fact]
    public void 그림이_움직이면_잡는다()
    {
        string b = Base();
        string o = CopyAndEdit(b, d =>
            Ws(d).DrawingsPart!.WorksheetDrawing.Elements<Xdr.TwoCellAnchor>().First()
                 .FromMarker!.RowId!.Text = "30");
        OutputVerifier.Verify(b, o, Doc).Should().ContainSingle().Which.Should().Contain("'그림 1' 위치가 바뀌었습니다");
    }

    [Fact]
    public void 그림_위치_항목으로_옮긴_그림은_정상()
    {
        string b = Base();
        string o = CopyAndEdit(b, d =>
            Ws(d).DrawingsPart!.WorksheetDrawing.Elements<Xdr.TwoCellAnchor>().ElementAt(1)
                 .FromMarker!.RowId!.Text = "30");
        OutputVerifier.Verify(b, o, Doc).Should().BeEmpty("글상자(순서 1)는 위치 항목에 들어 있다");
    }

    [Fact]
    public void 이미지가_없는_그림을_잡는다()
    {
        string b = Base();
        string o = CopyAndEdit(b, d => Ws(d).DrawingsPart!.WorksheetDrawing.Descendants<A.Blip>().First().Embed = "rIdNone");
        OutputVerifier.Verify(b, o, Doc).Should().Contain(p => p.Contains("빨간 X"));
    }

    [Fact]
    public void 구조_오류가_새로_생기면_잡는다()
    {
        string b = Base();
        // calcPr 을 통합문서 맨 끝(extLst 뒤)에 붙이는, 예전 버그와 같은 상황
        string o = CopyAndEdit(b, d =>
        {
            d.WorkbookPart!.Workbook.Append(new DocumentFormat.OpenXml.Spreadsheet.WorkbookExtensionList());
            d.WorkbookPart.Workbook.Append(new DocumentFormat.OpenXml.Spreadsheet.CalculationProperties());
        });
        OutputVerifier.Verify(b, o, Doc).Should().Contain(p => p.Contains("파일 구조 오류"));
    }

    [Fact]
    public void 열리지_않는_파일을_잡는다()
    {
        string b = Base();
        string o = _dir.File("broken.xlsx");
        File.WriteAllBytes(o, [1, 2, 3, 4]);
        OutputVerifier.Verify(b, o, Doc).Should().ContainSingle().Which.Should().Contain("열리지 않습니다");
    }

    [Fact]
    public void hwpx_정상과_망가진_본문()
    {
        string Make(string name, string body, bool mimeFirst = true)
        {
            string p = _dir.File(name);
            using var z = ZipFile.Open(p, ZipArchiveMode.Create);
            void Add(string n, string t, CompressionLevel l) { using var w = new StreamWriter(z.CreateEntry(n, l).Open()); w.Write(t); }
            if (mimeFirst) Add("mimetype", "application/hwp+zip", CompressionLevel.NoCompression);
            Add("Contents/section0.xml", body, CompressionLevel.Optimal);
            if (!mimeFirst) Add("mimetype", "application/hwp+zip", CompressionLevel.NoCompression);
            return p;
        }
        var doc = new DocumentConfig();
        OutputVerifier.Verify("", Make("ok.hwpx", "<a>1 &amp; 2</a>"), doc).Should().BeEmpty();
        OutputVerifier.Verify("", Make("bad.hwpx", "<a>1 & 2</a>"), doc).Should().Contain(p => p.Contains("망가졌습니다"));
        OutputVerifier.Verify("", Make("order.hwpx", "<a/>", mimeFirst: false), doc).Should().Contain(p => p.Contains("맨 앞"));
    }

    [Fact]
    public void 사진을_바꾼_소음측정일지도_통과()
    {
        XlsxBuilder.Create(_dir.File("noise.xlsx"), new SheetSpec("S") { Pictures = { new Pic("D12", "그림 1") } });
        string photo = _dir.File("p.jpg");
        File.WriteAllBytes(photo, XlsxBuilder.FakeJpeg);
        var cfg = new DocumentConfig
        {
            Id = "noise", Name = "소음", TemplateFile = "noise.xlsx", OutputPattern = "{dd} 소음.xlsx", SheetName = "S",
            Fields = [new() { Id = "ph", Type = FieldTypes.Photo, Cell = "D12" }],
        };
        var input = new DailyInput { Date = new DateTime(2026, 11, 5) };
        input.DocValues[DailyInput.Key("noise", "S", "ph")] = photo;
        var problems = new List<string>();

        GenericExcelWriter.Generate(input, cfg, _dir.Path, "output", (_, _, _) => XlsxBuilder.FakeJpeg, problems: problems);

        problems.Should().BeEmpty("사진 교체는 위치·개수를 바꾸지 않는다");
    }
}
