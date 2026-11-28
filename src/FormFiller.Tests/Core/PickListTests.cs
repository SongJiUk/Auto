using FluentAssertions;
using FormFiller.Core.Config;

namespace FormFiller.Tests.Core;

/// <summary>
/// [☰ 항목 선택] 목록: 여러 줄 항목(작업사항·장비 이름·장비 규격…)마다 따로 두고,
/// 추가·삭제·마지막으로 고른 것은 config.json 에 남아 다음 날에도 그대로 나온다.
/// </summary>
public class PickListTests
{
    private static AppConfig Cfg() => new() { PresetWorkItems = ["굴삭기 작업", "철거 작업"] };

    [Fact]
    public void Items_작업이_들어간_항목은_처음에_예전_작업목록으로_시작()
    {
        var c = Cfg();
        PickLists.Items(c, "금일 작업사항").Should().Equal("굴삭기 작업", "철거 작업");
    }

    [Fact]
    public void Items_작업이_아닌_항목은_빈_목록으로_시작()
    {
        var c = Cfg();
        PickLists.Items(c, "장비 이름").Should().BeEmpty();
        PickLists.Items(c, "장비 규격").Should().BeEmpty();
    }

    [Fact]
    public void Add_항목마다_따로_저장된다()
    {
        var c = Cfg();
        PickLists.Add(c, "장비 이름", "굴삭기 06W");
        PickLists.Add(c, "장비 규격", "0.6㎥");

        PickLists.Items(c, "장비 이름").Should().Equal("굴삭기 06W");
        PickLists.Items(c, "장비 규격").Should().Equal("0.6㎥");
        PickLists.Items(c, "금일 작업사항").Should().NotContain("굴삭기 06W");
    }

    [Fact]
    public void Add_같은_글자나_빈_글자는_넣지_않는다()
    {
        var c = Cfg();
        PickLists.Add(c, "장비 이름", " 굴삭기 ").Should().BeTrue();
        PickLists.Add(c, "장비 이름", "굴삭기").Should().BeFalse();
        PickLists.Add(c, "장비 이름", "  ").Should().BeFalse();
        PickLists.Items(c, "장비 이름").Should().Equal("굴삭기");
    }

    [Fact]
    public void Remove_목록과_마지막_고른것에서_모두_빠진다()
    {
        var c = Cfg();
        PickLists.SetLastChosen(c, "금일 작업사항", ["철거 작업"]);
        PickLists.Remove(c, "금일 작업사항", "철거 작업");

        PickLists.Items(c, "금일 작업사항").Should().Equal("굴삭기 작업");
        PickLists.LastChosen(c, "금일 작업사항").Should().BeEmpty();
    }

    [Fact]
    public void Remove_작업사항_목록을_다_지워도_예전_목록으로_되살아나지_않는다()
    {
        var c = Cfg();
        PickLists.Remove(c, "금일 작업사항", "굴삭기 작업");
        PickLists.Remove(c, "금일 작업사항", "철거 작업");
        PickLists.Items(c, "금일 작업사항").Should().BeEmpty();
    }

    [Fact]
    public void LastChosen_목록에_있는_것만_기억한다()
    {
        var c = Cfg();
        PickLists.SetLastChosen(c, "금일 작업사항", ["굴삭기 작업", "바로 넣은 글자"]);
        PickLists.LastChosen(c, "금일 작업사항").Should().Equal("굴삭기 작업");
        PickLists.LastChosen(c, "장비 이름").Should().BeEmpty();
    }

    [Fact]
    public void 저장하고_다시_읽어도_그대로()
    {
        var c = Cfg();
        PickLists.Add(c, "장비 이름", "굴삭기 06W");
        PickLists.SetLastChosen(c, "장비 이름", ["굴삭기 06W"]);

        var json = System.Text.Json.JsonSerializer.Serialize(c);
        var back = System.Text.Json.JsonSerializer.Deserialize<AppConfig>(json)!;

        PickLists.Items(back, "장비 이름").Should().Equal("굴삭기 06W");
        PickLists.LastChosen(back, "장비 이름").Should().Equal("굴삭기 06W");
    }
}
