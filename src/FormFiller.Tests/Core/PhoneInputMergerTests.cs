using FluentAssertions;
using FormFiller.Core.Models;
using FormFiller.Core.Phone;

namespace FormFiller.Tests.Core;

/// <summary>폰 입력을 지금 화면 값(DailyInput)에 합치는 규칙 테스트</summary>
public class PhoneInputMergerTests
{
    private static DailyInput Screen() => new()
    {
        Date    = new DateTime(2026, 11, 18),
        Weather = "흐림",
        TodayMorningWork = "- 기존 작업",
        Equipment =
        [
            new EquipmentEntry { Name = "B/H 380", Spec = "39.6t", PrevDayCount = 10, TodayCount = 0 },
            new EquipmentEntry { Name = "B/H 06W", PrevDayCount = 5, TodayCount = 1 },
        ],
        Personnel = [new PersonnelEntry { Role = "비계", PrevCumulativeCount = 30, TodayCount = 0 }],
        SujangjaePersonnel = [new PersonnelEntry { Role = "수장재", PrevCumulativeCount = 8 }],
    };

    private static PhoneInput Phone(
        IReadOnlyList<PhoneCount>? equipment = null,
        IReadOnlyList<PhoneCount>? personnel = null,
        string weather = "맑음",
        IReadOnlyList<string>? morning = null) => new()
    {
        Date      = new DateTime(2026, 11, 18),
        Weather   = weather,
        TodayMorningWork = morning ?? ["- 3층 철거"],
        Equipment = equipment ?? [],
        Personnel = personnel ?? [],
    };

    [Fact]
    public void Apply_MatchingNames_SetsTodayCountAndKeepsPrev()
    {
        var phone = Phone(equipment: [new PhoneCount("b/h380", "", 2)],          // 띄어쓰기·대소문자 무시
                          personnel: [new PhoneCount("비계", "", 3)]);

        var result = PhoneInputMerger.Apply(Screen(), phone);

        var eq = result.Input.Equipment.Single(e => e.Name == "B/H 380");
        eq.TodayCount.Should().Be(2);
        eq.PrevDayCount.Should().Be(10);
        eq.Spec.Should().Be("39.6t");
        result.Input.Personnel.Single().TodayCount.Should().Be(3);
        result.Unmatched.Should().BeEmpty();
    }

    [Fact]
    public void Apply_DoesNotMutateScreenInput()
    {
        var screen = Screen();

        PhoneInputMerger.Apply(screen, Phone(equipment: [new PhoneCount("B/H 380", "", 7)]));

        screen.Equipment[0].TodayCount.Should().Be(0);
        screen.Weather.Should().Be("흐림");
    }

    [Fact]
    public void Apply_NameNotOnScreen_AddsRowAndReportsIt()
    {
        var result = PhoneInputMerger.Apply(Screen(), Phone(equipment: [new PhoneCount("크레인", "25t", 1)]));

        var added = result.Input.Equipment.Single(e => e.Name == "크레인");
        added.TodayCount.Should().Be(1);
        added.Spec.Should().Be("25t");
        added.PrevDayCount.Should().Be(0);
        result.Unmatched.Should().ContainSingle().Which.Should().Contain("크레인");
    }

    [Fact]
    public void Apply_ScreenAlreadyHasDifferentValue_ReportsConflictAndUsesPhone()
    {
        // 화면에 B/H 06W 금일 1 이 이미 있는데 폰은 3
        var result = PhoneInputMerger.Apply(Screen(), Phone(equipment: [new PhoneCount("B/H 06W", "", 3)]));

        result.Input.Equipment.Single(e => e.Name == "B/H 06W").TodayCount.Should().Be(3);
        result.Conflicts.Should().ContainSingle().Which.Should().Contain("B/H 06W");
    }

    [Fact]
    public void Apply_RowNotSentByPhone_KeepsScreenValue()
    {
        var result = PhoneInputMerger.Apply(Screen(), Phone(equipment: [new PhoneCount("B/H 380", "", 2)]));

        result.Input.Equipment.Single(e => e.Name == "B/H 06W").TodayCount.Should().Be(1);
    }

    [Fact]
    public void Apply_SujangjaeRoles_GoToSujangjaeList()
    {
        var phone = Phone() with { SujangjaePersonnel = [new PhoneCount("수장재", "", 4)] };

        var result = PhoneInputMerger.Apply(Screen(), phone);

        result.Input.SujangjaePersonnel.Single().TodayCount.Should().Be(4);
        result.Input.SujangjaePersonnel.Single().PrevCumulativeCount.Should().Be(8);
        result.Input.Personnel.Single().TodayCount.Should().Be(0);
    }

    [Fact]
    public void Apply_WeatherAndWork_ReplacedOnlyWhenPhoneHasValue()
    {
        var filled = PhoneInputMerger.Apply(Screen(), Phone(weather: "맑음", morning: ["- 3층 철거", "- 폐기물 반출"]));
        filled.Input.Weather.Should().Be("맑음");
        filled.Input.TodayMorningWork.Should().Be("- 3층 철거\n- 폐기물 반출");

        var empty = PhoneInputMerger.Apply(Screen(), Phone(weather: "", morning: []));
        empty.Input.Weather.Should().Be("흐림");
        empty.Input.TodayMorningWork.Should().Be("- 기존 작업");
    }

    [Fact]
    public void Apply_DifferentDate_Throws()
    {
        var phone = Phone() with { Date = new DateTime(2026, 11, 17) };

        var act = () => PhoneInputMerger.Apply(Screen(), phone);

        act.Should().Throw<ArgumentException>();
    }
}
