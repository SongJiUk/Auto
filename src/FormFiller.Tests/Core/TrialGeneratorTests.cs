using System.Security.Cryptography;
using FluentAssertions;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Tests.Helpers;

namespace FormFiller.Tests.Core;

/// <summary>[🧪 시험 생성] — 저장된 설정 그대로, 임시 폴더에, 서식에서 새로 만든다.</summary>
public class TrialGeneratorTests : IDisposable
{
    private readonly TempDir _dir = new();
    private readonly string _template;

    public TrialGeneratorTests()
    {
        _template = DailyReportSample.CreateTemplate(_dir.File(DailyReportSample.TemplateName));
    }

    public void Dispose() => _dir.Dispose();

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));

    [Fact]
    public void 임시_폴더에_오늘_파일을_만들고_서식은_그대로()
    {
        var doc = DailyReportSample.Config();
        string before = Hash(_template);
        string folder = Path.Combine(_dir.Path, "trial");

        string path = TrialGenerator.Generate(doc, _template, folder, new DateTime(2026, 11, 5));

        File.Exists(path).Should().BeTrue();
        Path.GetDirectoryName(path).Should().Be(folder);
        Hash(_template).Should().Be(before);
    }

    [Fact]
    public void 넘겨준_설정은_바꾸지_않는다()
    {
        var doc = DailyReportSample.Config();
        doc.OutputPattern = "{yy}.{MM}.{dd} 시험";
        string templateBefore = doc.TemplateFile;

        TrialGenerator.Generate(doc, _template, Path.Combine(_dir.Path, "trial"), new DateTime(2026, 11, 5));

        doc.OutputPattern.Should().Be("{yy}.{MM}.{dd} 시험");
        doc.TemplateFile.Should().Be(templateBefore);
        doc.UsePreviousAsBase.Should().Be(DailyReportSample.Config().UsePreviousAsBase);
    }

    [Fact]
    public void 출력_이름이_비었거나_xlsx_가_아니면_xlsx_로()
    {
        var doc = DailyReportSample.Config();
        doc.OutputPattern = "";

        string path = TrialGenerator.Generate(doc, _template, Path.Combine(_dir.Path, "trial"), new DateTime(2026, 11, 5));

        Path.GetExtension(path).Should().Be(".xlsx");
    }
}
