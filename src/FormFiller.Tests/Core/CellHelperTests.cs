using FluentAssertions;
using FormFiller.Core.Excel;

namespace FormFiller.Tests.Core;

public class CellHelperTests
{
    [Theory]
    [InlineData("A1", "A", 1u)]
    [InlineData("B2", "B", 2u)]
    [InlineData("Z10", "Z", 10u)]
    [InlineData("AA3", "AA", 3u)]
    public void ParseAddress_ReturnsColumnAndRow_Correctly(string address, string expectedCol, uint expectedRow)
    {
        var (col, row) = CellHelper.ParseAddress(address);

        col.Should().Be(expectedCol);
        row.Should().Be(expectedRow);
    }

    [Theory]
    [InlineData("A", 1u)]
    [InlineData("B", 2u)]
    [InlineData("Z", 26u)]
    [InlineData("AA", 27u)]
    [InlineData("AB", 28u)]
    public void ColLetterToIndex_ReturnsCorrectIndex(string col, uint expected)
    {
        CellHelper.ColLetterToIndex(col).Should().Be(expected);
    }
}
