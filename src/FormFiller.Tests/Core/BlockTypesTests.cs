using FluentAssertions;
using FormFiller.Core.Config;

namespace FormFiller.Tests.Core;

/// <summary>설정 화면 문서 카드에 보이는 블록 종류 이름·칸 요약 (2026-11-22: 여러 줄·표에 빈 주소 칸이 보이던 문제)</summary>
public class BlockTypesTests
{
    [Theory]
    [InlineData(BlockTypes.Lines, "여러 줄")]
    [InlineData(BlockTypes.Table, "표")]
    [InlineData(BlockTypes.Formula, "계산식")]
    [InlineData("equipment_table", "equipment_table")]
    public void TypeTitle_종류이름_한글로_모르는종류는_그대로(string type, string expected) =>
        BlockTypes.TypeTitle(type).Should().Be(expected);

    [Fact]
    public void CellSummary_여러줄은_첫칸부터_끝칸과_칸수() =>
        BlockTypes.CellSummary(new FieldConfig { Type = BlockTypes.Lines, Cells = ["B10", "B11", "B12"] })
            .Should().Be("B10 ~ B12 (3칸)");

    [Fact]
    public void CellSummary_한칸이면_주소만() =>
        BlockTypes.CellSummary(new FieldConfig { Type = BlockTypes.Lines, Cells = ["B10"] }).Should().Be("B10");

    [Fact]
    public void CellSummary_표는_행범위() =>
        BlockTypes.CellSummary(new FieldConfig { Type = BlockTypes.Table, StartRow = 20, EndRow = 32 })
            .Should().Be("20 ~ 32행");

    [Fact]
    public void CellSummary_사진구역은_칸개수() =>
        BlockTypes.CellSummary(new FieldConfig
        {
            Type = BlockTypes.Photos,
            PhotoAreas = [new PhotoArea { Range = "A4:H15" }, new PhotoArea { Range = "A17:H28" }],
        }).Should().Be("사진 칸 2개");

    [Fact]
    public void CellSummary_칸이_없으면_입력만() =>
        BlockTypes.CellSummary(new FieldConfig { Type = BlockTypes.Value }).Should().Be("칸 없음 (입력만)");
}
