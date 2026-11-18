using System.IO.Compression;
using System.Text;
using FluentAssertions;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Core.Models;
using FormFiller.Tests.Helpers;

namespace FormFiller.Tests.Core;

/// <summary>한글(hwpx) 자리 글자 바꾸기 테스트. hwpx 는 zip 안의 XML 이다.</summary>
public class HwpxWriterTests : IDisposable
{
    private readonly TempDir _dir = new();
    public void Dispose() => _dir.Dispose();

    private string CreateHwpx(string body)
    {
        string path = _dir.File("plan.hwpx");
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        void Add(string name, string text, CompressionLevel level)
        {
            using var w = new StreamWriter(zip.CreateEntry(name, level).Open(), new UTF8Encoding(false));
            w.Write(text);
        }
        Add("mimetype", "application/hwp+zip", CompressionLevel.NoCompression);
        Add("Contents/section0.xml", $"<hs:sec xmlns:hs=\"x\" xmlns:hp=\"y\"><hp:p><hp:t>{body}</hp:t></hp:p></hs:sec>", CompressionLevel.Optimal);
        Add("Contents/header.xml", "<hh:head xmlns:hh=\"z\"/>", CompressionLevel.Optimal);
        Add("BinData/image1.png", "PNGDATA", CompressionLevel.Optimal);
        return path;
    }

    private static DocumentConfig Plan() => new()
    {
        Id = "plan", Name = "작업계획서", TemplateFile = "plan.hwpx", OutputPattern = "{yy}.{MM}.{dd} 계획서.hwpx",
        Fields =
        [
            new() { Id = "date", Label = "날짜", Type = FieldTypes.DateKorean },
            new() { Id = "weather", Label = "날씨", Type = FieldTypes.Text },
            new() { Id = "today_work", Label = "작업", Type = FieldTypes.Multiline, Placeholder = "{{금일작업}}" },
        ],
    };

    private static string ReadEntry(string path, string name)
    {
        using var zip = ZipFile.OpenRead(path);
        using var r = new StreamReader(zip.GetEntry(name)!.Open());
        return r.ReadToEnd();
    }

    [Fact]
    public void 자리_글자를_값으로_바꾸고_나머지는_그대로()
    {
        string tpl = CreateHwpx("작성일: {{날짜}} / 날씨: {{날씨}} / {{금일작업}}");
        var input = new DailyInput { Date = new DateTime(2026, 11, 5), Weather = "비 & 바람", TodayMorningWork = "철거", TodayAfternoonWork = "반출 <2차>" };
        string outPath = _dir.File("out.hwpx");

        HwpxWriter.Generate(tpl, outPath, Plan(), input);

        string xml = ReadEntry(outPath, "Contents/section0.xml");
        xml.Should().Contain("작성일: 2026년 11월 05일 목요일");
        xml.Should().Contain("날씨: 비 &amp; 바람", "XML 특수문자는 바꿔 넣는다");
        xml.Should().Contain("철거<hp:lineBreak/>반출 &lt;2차&gt;", "줄바꿈은 한글 줄 나눔으로");
        xml.Should().NotContain("{{");
        ReadEntry(outPath, "BinData/image1.png").Should().Be("PNGDATA");
        ReadEntry(outPath, "Contents/header.xml").Should().Be("<hh:head xmlns:hh=\"z\"/>");
    }

    [Fact]
    public void mimetype은_맨_앞에_압축하지_않고_둔다()
    {
        string tpl = CreateHwpx("{{날짜}}");
        string outPath = _dir.File("out.hwpx");
        HwpxWriter.Generate(tpl, outPath, Plan(), new DailyInput { Date = DateTime.Today });

        using var zip = ZipFile.OpenRead(outPath);
        var first = zip.Entries[0];
        first.FullName.Should().Be("mimetype");
        first.CompressedLength.Should().Be(first.Length, "압축 안 함");
        zip.Entries.Select(e => e.FullName).Should().BeEquivalentTo(
            "mimetype", "Contents/section0.xml", "Contents/header.xml", "BinData/image1.png");
    }

    [Fact]
    public void 서식에_없는_자리_글자는_알린다()
    {
        string tpl = CreateHwpx("{{날짜}}");
        var notes = new List<string>();
        HwpxWriter.Generate(tpl, _dir.File("out.hwpx"), Plan(), new DailyInput { Date = DateTime.Today }, notes);
        notes.Should().Contain(n => n.Contains("{{날씨}}")).And.Contain(n => n.Contains("{{금일작업}}"))
             .And.NotContain(n => n.Contains("{{날짜}}"));
    }

    [Fact]
    public void 서식의_자리_글자_목록을_읽는다()
    {
        string tpl = CreateHwpx("{{날짜}} {{날씨}} {{날짜}} {not} {{}}");
        HwpxWriter.FindPlaceholders(tpl).Should().Equal("{{날짜}}", "{{날씨}}");
    }

    [Fact]
    public void 원본_서식은_바뀌지_않는다()
    {
        string tpl = CreateHwpx("{{날짜}}");
        byte[] before = File.ReadAllBytes(tpl);
        HwpxWriter.Generate(tpl, _dir.File("out.hwpx"), Plan(), new DailyInput { Date = DateTime.Today });
        File.ReadAllBytes(tpl).Should().Equal(before);
    }

    [Fact]
    public void 엔진은_hwpx_서식이면_항상_서식에서_만든다()
    {
        CreateHwpx("{{날짜}}");
        var cfg = Plan();
        string thu = GenericExcelWriter.Generate(new DailyInput { Date = new DateTime(2026, 11, 5) }, cfg, _dir.Path, "output");
        string fri = GenericExcelWriter.Generate(new DailyInput { Date = new DateTime(2026, 11, 6) }, cfg, _dir.Path, "output");

        ReadEntry(thu, "Contents/section0.xml").Should().Contain("11월 05일");
        ReadEntry(fri, "Contents/section0.xml").Should().Contain("11월 06일", "앞 날짜 파일엔 자리 글자가 없으니 서식을 써야 한다");
    }
}
