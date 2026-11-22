using FluentAssertions;
using FormFiller.Core.Config;
using FormFiller.Tests.Helpers;

namespace FormFiller.Tests.Core;

/// <summary>
/// 화면 정리 4단계: 모든 PC 가 같은 설정(드라이브의 config.json)을 쓰는지 띠로 보여 주고,
/// 다른 PC 에서 설정을 바꿨으면 켤 때 한 줄 알린다.
/// </summary>
public class SharedConfigTests : IDisposable
{
    private const string DriveDir = @"G:\내 드라이브\업무자동화\";
    private static readonly DateTime Saved = new(2026, 11, 23, 9, 12, 0);

    private readonly TempDir _dir = new();
    public void Dispose() => _dir.Dispose();

    private static AppConfig Cfg(string? by, DateTime? at) => new() { LastSavedBy = by ?? "", LastSavedAt = at };

    [Fact]
    public void 드라이브_안이면_공유_폴더와_마지막_저장_PC_시각()
    {
        var s = SharedConfig.Describe(Cfg("OFFICE-PC", Saved), DriveDir, "HOME-PC", lastSeen: null);

        s.IsShared.Should().BeTrue();
        s.Text.Should().Contain("공유 설정").And.Contain(@"내 드라이브\업무자동화")
              .And.Contain("OFFICE-PC").And.Contain("11/23 09:12");
    }

    [Fact]
    public void 이_PC가_저장했으면_이_PC로_표시() =>
        SharedConfig.Describe(Cfg("HOME-PC", Saved), DriveDir, "HOME-PC", null).Text.Should().Contain("이 PC");

    [Fact]
    public void 저장_기록이_없으면_그렇게_표시() =>
        SharedConfig.Describe(Cfg(null, null), DriveDir, "HOME-PC", null).Text.Should().Contain("기록 없음");

    [Fact]
    public void 드라이브_밖이면_경고()
    {
        var s = SharedConfig.Describe(Cfg("HOME-PC", Saved), @"C:\Users\me\Desktop\업무자동화\", "HOME-PC", null);

        s.IsShared.Should().BeFalse();
        s.Text.Should().StartWith("⚠").And.Contain("공유되지 않");
    }

    [Fact]
    public void 다른_PC가_마지막으로_본_뒤에_바꿨으면_알림()
    {
        var s = SharedConfig.Describe(Cfg("OFFICE-PC", Saved), DriveDir, "HOME-PC", lastSeen: Saved.AddHours(-1));
        s.Notice.Should().Contain("OFFICE-PC").And.Contain("11/23 09:12");
    }

    [Fact]
    public void 이미_본_변경이나_내가_바꾼_것이나_처음_켠_PC는_알림_없음()
    {
        SharedConfig.Describe(Cfg("OFFICE-PC", Saved), DriveDir, "HOME-PC", lastSeen: Saved).Notice.Should().BeNull();
        SharedConfig.Describe(Cfg("HOME-PC", Saved), DriveDir, "HOME-PC", lastSeen: Saved.AddDays(-1)).Notice.Should().BeNull();
        SharedConfig.Describe(Cfg("OFFICE-PC", Saved), DriveDir, "HOME-PC", lastSeen: null).Notice.Should().BeNull();
    }

    [Fact]
    public void 본_시각은_PC별_파일에_저장하고_다시_읽는다()
    {
        string path = _dir.File(@"seen\config_seen.txt");
        SharedConfig.ReadSeen(path).Should().BeNull();

        SharedConfig.WriteSeen(path, Saved);
        SharedConfig.ReadSeen(path).Should().Be(Saved);
    }

    [Fact]
    public void 본_시각_파일이_깨져_있으면_없는_것으로()
    {
        string path = _dir.File("config_seen.txt");
        File.WriteAllText(path, "엉망");
        SharedConfig.ReadSeen(path).Should().BeNull();
    }

    [Fact]
    public void 저장하면_PC이름과_시각을_남긴다()
    {
        var cfg = new AppConfig();
        SharedConfig.Stamp(cfg, "HOME-PC", Saved);
        cfg.LastSavedBy.Should().Be("HOME-PC");
        cfg.LastSavedAt.Should().Be(Saved);
    }
}
