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

    [Fact]
    public void 칸_안_글자_inlineStr_를_읽고_덮어쓴다()
    {
        // openpyxl 등 다른 프로그램이 만든 xlsx 는 공유 문자열 대신 칸 안에 글자를 넣는다
        using var dir = new FormFiller.Tests.Helpers.TempDir();
        string path = dir.File("inline.xlsx");
        FormFiller.Tests.Helpers.XlsxBuilder.Create(path, new FormFiller.Tests.Helpers.SheetSpec("S") { ["A1"] = 1, ["B1"] = 2 });
        using (var d = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, true))
        {
            var ws = CellHelper.FindWorksheetPart(d.WorkbookPart!, "S")!;
            foreach (var c in ws.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>())
            {
                c.CellValue = null;
                c.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString;
                c.InlineString = new DocumentFormat.OpenXml.Spreadsheet.InlineString(
                    new DocumentFormat.OpenXml.Spreadsheet.Text(c.CellReference == "A1" ? "B/H 380" : "덤프"));
            }
            CellHelper.GetCellValue(ws, "A1").Should().Be("B/H 380");

            CellHelper.SetCellString(d, ws, "A1", "새 이름");
            CellHelper.SetCellInt(ws, "B1", 7);
            CellHelper.GetCellValue(ws, "A1").Should().Be("새 이름");
            CellHelper.GetCellInt(ws, "B1").Should().Be(7);
        }
        using var r = new FormFiller.Tests.Helpers.XlsxReader(path);
        r.ValidationErrors().Should().BeEmpty("옛 칸 안 글자가 남으면 구조 오류");
    }
}
