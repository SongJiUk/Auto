using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace FormFiller.Tests.Helpers;

/// <summary>수식 칸: 수식과 저장된 값(캐시)</summary>
public record F(string Formula, string? Cached = null);

/// <summary>그림(사진) 하나: 왼쪽 위 칸, 이름, 크기(EMU)</summary>
public record Pic(string Cell, string Name, long Cx = 1_905_000, long Cy = 1_428_750);

/// <summary>글상자 하나 (그림 위치 테스트용)</summary>
public record Box(string Cell, string Name);

/// <summary>
/// 테스트용 시트 하나. Cells 의 값: string → 공유 문자열, int/double → 숫자, F → 수식.
/// </summary>
public class SheetSpec(string name)
{
    public string Name { get; } = name;
    public Dictionary<string, object> Cells { get; } = [];
    public List<Pic> Pictures { get; } = [];
    public List<Box> Boxes { get; } = [];

    public object this[string cell] { set => Cells[cell] = value; }
}

/// <summary>
/// 실제 서식 대신 쓰는 가짜 xlsx 를 만든다 (실제 서식은 개인정보라 저장소에 없다).
/// 엑셀이 저장하는 모양(공유 문자열, 수식 + 캐시값, 그림 앵커)을 흉내 낸다.
/// </summary>
public static class XlsxBuilder
{
    /// <summary>1×1 PNG (그림 칸의 빈 이미지)</summary>
    public static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    /// <summary>엔진이 사진을 줄일 때 쓰는 함수 대신, 받은 크기를 기록하고 고정 JPEG 바이트를 돌려준다.</summary>
    public static readonly byte[] FakeJpeg = [0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00, 0xFF, 0xD9];

    public static string Create(string path, params SheetSpec[] sheets)
    {
        using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var wb = doc.AddWorkbookPart();
        wb.Workbook = new Workbook(new Sheets());
        var sst = wb.AddNewPart<SharedStringTablePart>();
        sst.SharedStringTable = new SharedStringTable();
        var strings = new List<string>();

        uint sheetId = 1;
        foreach (var spec in sheets)
        {
            var ws = wb.AddNewPart<WorksheetPart>();
            var data = new SheetData();
            foreach (var rowGroup in spec.Cells
                         .Select(kv => (Addr: kv.Key, Val: kv.Value, Pos: Parse(kv.Key)))
                         .GroupBy(x => x.Pos.Row).OrderBy(g => g.Key))
            {
                var row = new Row { RowIndex = rowGroup.Key };
                foreach (var c in rowGroup.OrderBy(x => x.Pos.Col))
                    row.Append(MakeCell(c.Addr, c.Val, strings));
                data.Append(row);
            }
            ws.Worksheet = new Worksheet(data);

            if (spec.Pictures.Count > 0 || spec.Boxes.Count > 0)
                AddDrawing(ws, spec);

            ws.Worksheet.Save();
            wb.Workbook.Sheets!.Append(new Sheet { Id = wb.GetIdOfPart(ws), SheetId = sheetId++, Name = spec.Name });
        }

        foreach (var s in strings) sst.SharedStringTable.Append(new SharedStringItem(new Text(s)));
        sst.SharedStringTable.Count = (uint)strings.Count;
        sst.SharedStringTable.UniqueCount = (uint)strings.Count;
        wb.Workbook.Save();
        return path;
    }

    private static Cell MakeCell(string addr, object val, List<string> strings)
    {
        var cell = new Cell { CellReference = addr };
        switch (val)
        {
            case string s:
                int idx = strings.IndexOf(s);
                if (idx < 0) { strings.Add(s); idx = strings.Count - 1; }
                cell.DataType  = CellValues.SharedString;
                cell.CellValue = new CellValue(idx.ToString());
                break;
            case int or double:
                cell.CellValue = new CellValue(Convert.ToString(val, System.Globalization.CultureInfo.InvariantCulture)!);
                break;
            case F f:
                cell.CellFormula = new CellFormula(f.Formula);
                if (f.Cached != null) cell.CellValue = new CellValue(f.Cached);
                break;
        }
        return cell;
    }

    private static void AddDrawing(WorksheetPart ws, SheetSpec spec)
    {
        var dp = ws.AddNewPart<DrawingsPart>();
        var sb = new System.Text.StringBuilder();
        sb.Append("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" ")
          .Append("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" ")
          .Append("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

        int id = 2;
        foreach (var p in spec.Pictures)
        {
            var img = dp.AddImagePart(ImagePartType.Png);
            using (var ms = new MemoryStream(TinyPng)) img.FeedData(ms);
            string rid = dp.GetIdOfPart(img);
            var (c, r) = Parse(p.Cell);
            sb.Append(Anchor((int)c - 1, (int)r - 1,
                $"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"{id++}\" name=\"{p.Name}\"/><xdr:cNvPicPr/></xdr:nvPicPr>" +
                $"<xdr:blipFill><a:blip r:embed=\"{rid}\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>" +
                $"<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{p.Cx}\" cy=\"{p.Cy}\"/></a:xfrm>" +
                "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>"));
        }
        foreach (var b in spec.Boxes)
        {
            var (c, r) = Parse(b.Cell);
            sb.Append(Anchor((int)c - 1, (int)r - 1,
                $"<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"{id++}\" name=\"{b.Name}\"/><xdr:cNvSpPr txBox=\"1\"/></xdr:nvSpPr>" +
                "<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"952500\" cy=\"476250\"/></a:xfrm>" +
                "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>"));
        }
        sb.Append("</xdr:wsDr>");

        dp.WorksheetDrawing = new Xdr.WorksheetDrawing(sb.ToString());
        dp.WorksheetDrawing.Save();
        ws.Worksheet.Append(new Drawing { Id = ws.GetIdOfPart(dp) });
    }

    private static string Anchor(int col, int row, string body) =>
        "<xdr:twoCellAnchor>" +
        $"<xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
        $"<xdr:to><xdr:col>{col + 2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row + 5}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" +
        body + "<xdr:clientData/></xdr:twoCellAnchor>";

    /// <summary>"AB12" → (열 28, 행 12)</summary>
    public static (uint Col, uint Row) Parse(string addr)
    {
        int i = 0;
        uint col = 0;
        while (i < addr.Length && char.IsLetter(addr[i])) col = col * 26 + (uint)(char.ToUpperInvariant(addr[i++]) - 'A' + 1);
        return (col, uint.Parse(addr[i..]));
    }
}

/// <summary>만든 파일을 열어 값을 확인하는 도우미</summary>
public sealed class XlsxReader : IDisposable
{
    private readonly SpreadsheetDocument _doc;
    public XlsxReader(string path) => _doc = SpreadsheetDocument.Open(path, false);

    public WorkbookPart Workbook => _doc.WorkbookPart!;

    public WorksheetPart Sheet(string name) =>
        FormFiller.Core.Excel.CellHelper.FindWorksheetPart(Workbook, name)
        ?? throw new InvalidOperationException($"시트 없음: {name}");

    public string? Value(string sheet, string cell) =>
        FormFiller.Core.Excel.CellHelper.GetCellValue(Sheet(sheet), cell);

    public int? Int(string sheet, string cell) =>
        FormFiller.Core.Excel.CellHelper.GetCellInt(Sheet(sheet), cell);

    public Cell? RawCell(string sheet, string cell) =>
        Sheet(sheet).Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference?.Value == cell);

    /// <summary>Open XML 스키마 검사 오류 (엑셀이 "파일 손상"이라고 할 만한 것)</summary>
    public List<string> ValidationErrors() =>
        new DocumentFormat.OpenXml.Validation.OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(_doc)
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}")
            .ToList();

    public List<Xdr.TwoCellAnchor> Anchors(string sheet) =>
        Sheet(sheet).DrawingsPart?.WorksheetDrawing?.Elements<Xdr.TwoCellAnchor>().ToList() ?? [];

    public void Dispose() => _doc.Dispose();
}

/// <summary>테스트마다 쓰고 지우는 임시 폴더 (프로그램 폴더 역할)</summary>
public sealed class TempDir : IDisposable
{
    public string Path { get; } = Directory.CreateTempSubdirectory("ff_test_").FullName;
    public string File(string name) => System.IO.Path.Combine(Path, name);
    public void Dispose() { try { Directory.Delete(Path, true); } catch { } }
}
