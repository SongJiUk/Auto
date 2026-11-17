using FluentAssertions;
using FormFiller.Core.Config;

namespace FormFiller.Tests.Core;

public class GoogleDriveTests
{
    [Theory]
    [InlineData(@"G:\내 드라이브\FormFiller\사진", @"FormFiller\사진")]
    [InlineData(@"G:\내 드라이브", "")]
    [InlineData(@"G:\My Drive\FormFiller", "FormFiller")]
    [InlineData(@"C:\Users\kim\my drive\a\b", @"a\b")]      // 대소문자 무시
    public void RelativeToDrive_드라이브_안_경로는_뒤쪽만_돌려준다(string path, string expected)
        => GoogleDrive.RelativeToDrive(path).Should().Be(expected);

    [Theory]
    [InlineData(@"C:\Users\kim\Desktop\FormFiller")]
    [InlineData(@"G:\내 드라이브2\FormFiller")]                 // 비슷한 이름은 아님
    [InlineData("")]
    [InlineData(null)]
    public void RelativeToDrive_드라이브_밖이면_null(string? path)
        => GoogleDrive.RelativeToDrive(path).Should().BeNull();

    [Fact]
    public void Remap_경로가_있으면_그대로()
    {
        var result = GoogleDrive.Remap(@"G:\내 드라이브\사진", [@"H:\내 드라이브"], _ => true);
        result.Should().Be(@"G:\내 드라이브\사진");
    }

    [Fact]
    public void Remap_드라이브_문자가_다르면_이_PC_드라이브로_바꾼다()
    {
        var existing = new HashSet<string> { @"H:\내 드라이브\FormFiller\사진" };
        var result = GoogleDrive.Remap(@"G:\내 드라이브\FormFiller\사진", [@"H:\내 드라이브"], existing.Contains);
        result.Should().Be(@"H:\내 드라이브\FormFiller\사진");
    }

    [Fact]
    public void Remap_한글_영문_드라이브_이름이_달라도_찾는다()
    {
        var existing = new HashSet<string> { @"G:\My Drive\FormFiller" };
        var result = GoogleDrive.Remap(@"G:\내 드라이브\FormFiller", [@"G:\My Drive"], existing.Contains);
        result.Should().Be(@"G:\My Drive\FormFiller");
    }

    [Fact]
    public void Remap_어디에도_없으면_원래_경로()
    {
        var result = GoogleDrive.Remap(@"G:\내 드라이브\없는폴더", [@"H:\내 드라이브"], _ => false);
        result.Should().Be(@"G:\내 드라이브\없는폴더");
    }

    [Fact]
    public void Remap_드라이브_밖_경로는_바꾸지_않는다()
    {
        var result = GoogleDrive.Remap(@"D:\사진", [@"H:\내 드라이브"], _ => false);
        result.Should().Be(@"D:\사진");
    }

    [Fact]
    public void DriveChecker_사진칸_있는_문서만_사진폴더_점검()
    {
        var dir = Directory.CreateTempSubdirectory("ff_drive_").FullName;
        try
        {
            var cfg = new AppConfig
            {
                OutputFolder = "output",
                Documents =
                [
                    new DocumentConfig { Name = "소음", Fields = [new FieldConfig { Type = FieldTypes.Photo }] },
                    new DocumentConfig { Name = "공사일보", Fields = [new FieldConfig { Type = FieldTypes.Text }] },
                ],
            };
            var items = DriveChecker.Run(cfg, dir);

            items.Select(i => i.Title).Should().Contain("사진 폴더 · 소음").And.NotContain("사진 폴더 · 공사일보");
            // 출력 폴더가 아직 없으면 경고(처음 만들 때 생김)이지 오류가 아니다
            items.Single(i => i.Title == "출력 폴더").State.Should().Be(CheckState.Warning);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void DriveChecker_쓸_수_있는_드라이브_밖_출력폴더는_경고()
    {
        var dir = Directory.CreateTempSubdirectory("ff_drive_").FullName;
        try
        {
            Directory.CreateDirectory(Path.Combine(dir, "output"));
            var items = DriveChecker.Run(new AppConfig { OutputFolder = "output" }, dir);
            var output = items.Single(i => i.Title == "출력 폴더");
            output.State.Should().Be(CheckState.Warning);
            output.Detail.Should().Contain("드라이브 밖");
            Directory.GetFiles(Path.Combine(dir, "output")).Should().BeEmpty();   // 쓰기 시험 파일은 지운다
        }
        finally { Directory.Delete(dir, true); }
    }
}
