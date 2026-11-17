using FluentAssertions;
using FormFiller.Core.Phone;

namespace FormFiller.Tests.Core;

/// <summary>폰 사진의 라벨로 PC 사진 칸을 찾는 규칙 + inbox 위치 계산 테스트</summary>
public class PhonePhotoMatcherTests
{
    private static PhonePhoto Photo(string label, string file = "p.jpg") => new(file, label, "", @"C:\inbox\" + file);

    [Fact]
    public void Match_ExactLabel_WinsOverPartial()
    {
        // "장비 사진" 칸이 먼저 있어도, 이름이 똑같은 "장비" 칸에 넣는다
        var result = PhonePhotoMatcher.Match(["장비 사진", "장비"], [Photo("장비")]);

        result.Placed.Should().ContainSingle().Which.TargetIndex.Should().Be(1);
        result.Unplaced.Should().BeEmpty();
    }

    [Fact]
    public void Match_PartialLabel_IgnoresSpaceAndCase()
    {
        var result = PhonePhotoMatcher.Match(["현장 전경", "장비 사진"], [Photo("장비사진")]);

        result.Placed.Should().ContainSingle().Which.TargetIndex.Should().Be(1);
    }

    [Fact]
    public void Match_SameLabelTwice_FillsNextTarget()
    {
        var result = PhonePhotoMatcher.Match(["측정사진", "측정사진"], [Photo("측정사진", "a.jpg"), Photo("측정사진", "b.jpg")]);

        result.Placed.Select(p => (p.TargetIndex, p.Photo.File)).Should().Equal((0, "a.jpg"), (1, "b.jpg"));
    }

    [Fact]
    public void Match_NoMatchingOrEmptyLabel_IsUnplaced()
    {
        var result = PhonePhotoMatcher.Match(["장비"], [Photo("전경"), Photo("")]);

        result.Placed.Should().BeEmpty();
        result.Unplaced.Should().HaveCount(2);
    }

    [Fact]
    public void Match_MoreSamePhotosThanTargets_ExtraIsUnplaced()
    {
        var result = PhonePhotoMatcher.Match(["장비"], [Photo("장비", "a.jpg"), Photo("장비", "b.jpg")]);

        result.Placed.Should().ContainSingle().Which.Photo.File.Should().Be("a.jpg");
        result.Unplaced.Should().ContainSingle().Which.File.Should().Be("b.jpg");
    }

    [Fact]
    public void ResolveInboxFolder_Empty_UsesSiblingOfExeFolder()
    {
        string baseDir = Path.Combine(Path.GetTempPath(), "drive", "FormFiller");

        PhoneInbox.ResolveInboxFolder("", baseDir)
                  .Should().Be(Path.GetFullPath(Path.Combine(Path.GetTempPath(), "drive", "FormFiller_inbox")));
    }

    [Fact]
    public void ResolveInboxFolder_AbsolutePath_UsedAsIs()
    {
        string abs = Path.Combine(Path.GetTempPath(), "my_inbox");

        PhoneInbox.ResolveInboxFolder(abs, @"C:\anything").Should().Be(Path.GetFullPath(abs));
    }
}
