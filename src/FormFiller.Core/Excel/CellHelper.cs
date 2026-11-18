using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FormFiller.Core.Excel;

/// <summary>
/// Open XML SDK로 엑셀 셀을 읽고 쓰는 저수준 헬퍼 모음.
///
/// [Open XML 핵심 개념]
/// - SpreadsheetDocument → WorkbookPart → WorksheetPart → Worksheet → SheetData → Row → Cell
/// - 문자열 값은 파일 크기 절약을 위해 SharedStringTable에 한 번만 저장되고,
///   Cell.DataType = CellValues.SharedString, Cell.CellValue = "인덱스"(정수 문자열) 형태로 참조된다.
/// - 숫자는 Cell.DataType이 없거나(null) CellValues.Number이고, CellValue에 직접 숫자 문자열이 들어간다.
/// - 셀 주소는 "C6", "AD6" 등 열(문자)+ 행(숫자) 형식이다.
/// </summary>
public static class CellHelper
{
    // ── 셀 주소 파싱 ──────────────────────────────────────────────────

    /// <summary>
    /// "AD6" → 열 문자열("AD"), 행 번호(6)로 분리한다.
    /// </summary>
    public static (string col, uint row) ParseAddress(string cellAddress)
    {
        // 앞쪽 알파벳 = 열, 뒤쪽 숫자 = 행
        int i = 0;
        while (i < cellAddress.Length && char.IsLetter(cellAddress[i]))
            i++;

        string col = cellAddress[..i].ToUpperInvariant();
        uint   row = uint.Parse(cellAddress[i..]);
        return (col, row);
    }

    /// <summary>
    /// 열 문자열("AD") → 1-based 열 인덱스(30).
    /// Open XML은 셀 삽입 순서를 열 인덱스로 비교하므로 필요하다.
    /// </summary>
    public static uint ColLetterToIndex(string col)
    {
        // A=1, B=2, … Z=26, AA=27, AB=28 …
        uint index = 0;
        foreach (char c in col)
            index = index * 26 + (uint)(c - 'A' + 1);
        return index;
    }

    // ── 셀 읽기 ───────────────────────────────────────────────────────

    /// <summary>
    /// 셀의 텍스트 값을 반환한다 (SharedString·인라인 문자열·숫자 모두 처리).
    /// 셀이 없으면 null 반환.
    /// </summary>
    /// <summary>
    /// 셀에 수식이 있는지. 수식 셀은 값을 덮어쓰지 않는다 (예: 2·3번째 시트 날짜가 첫 시트를 참조하는 경우).
    /// </summary>
    public static bool HasFormula(WorksheetPart wsPart, string cellAddress)
        => FindCell(wsPart.Worksheet, cellAddress)?.CellFormula != null;

    public static string? GetCellValue(WorksheetPart wsPart, string cellAddress)
    {
        var cell = FindCell(wsPart.Worksheet, cellAddress);
        if (cell == null) return null;

        return ResolveCellValue(wsPart.OpenXmlPackage as SpreadsheetDocument, cell);
    }

    /// <summary>
    /// 셀의 숫자 값을 반환한다. 파싱 실패 시 null.
    /// </summary>
    public static int? GetCellInt(WorksheetPart wsPart, string cellAddress)
    {
        string? raw = GetCellValue(wsPart, cellAddress);
        if (raw == null) return null;

        // 소수점이 있는 "3.0" 같은 형태도 처리
        if (double.TryParse(raw, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out double d))
            return (int)d;

        return null;
    }

    // ── 셀 쓰기: 문자열 ───────────────────────────────────────────────

    /// <summary>
    /// 셀에 문자열 값을 쓴다. SharedStringTable을 통해 저장한다.
    ///
    /// [왜 SharedString을 써야 하나?]
    /// 엑셀이 파일을 열 때 t="s"(SharedString) 타입이 아닌 인라인 문자열(t="str" 또는 무타입)을
    /// 일부 버전에서 올바르게 렌더링하지 못하거나, 기존 서식·병합을 망가뜨릴 수 있다.
    /// 따라서 원본 셀이 SharedString이면 SharedString으로, 아니면 인라인으로 덮는다.
    /// </summary>
    public static void SetCellString(SpreadsheetDocument doc, WorksheetPart wsPart,
                                     string cellAddress, string value)
    {
        var cell = GetOrCreateCell(wsPart.Worksheet, cellAddress);

        // 원본이 SharedString 타입이었으면 SharedString으로 저장 (서식 보존 우선)
        bool wasSharedString = cell.DataType?.Value == CellValues.SharedString;

        if (wasSharedString)
        {
            int idx = GetOrAddSharedString(doc, value);
            cell.DataType  = CellValues.SharedString;
            cell.CellValue = new CellValue(idx.ToString());
        }
        else
        {
            // 인라인 문자열(t="str")로 저장. 칸 안 글자(<is>)가 남아 있으면 지운다 (두 값이 함께 있으면 구조 오류)
            cell.InlineString?.Remove();
            cell.DataType  = CellValues.String;
            cell.CellValue = new CellValue(value);
        }
    }

    // ── 셀 쓰기: 숫자 ─────────────────────────────────────────────────

    /// <summary>
    /// 셀에 정수 값을 쓴다.
    /// DataType을 null(기본 숫자 타입)로 두고 CellValue에 숫자 문자열을 넣는다.
    /// </summary>
    public static void SetCellInt(WorksheetPart wsPart, string cellAddress, int value)
    {
        var cell = GetOrCreateCell(wsPart.Worksheet, cellAddress);
        cell.InlineString?.Remove();                        // 글자 칸이었으면 남은 글자를 지운다
        cell.DataType  = null;                              // 숫자는 DataType 없음
        cell.CellValue = new CellValue(value.ToString());
    }

    // ── fullCalcOnLoad 설정 ───────────────────────────────────────────

    /// <summary>
    /// 파일을 열 때 모든 수식을 재계산하도록 플래그를 설정한다.
    /// J열(누계), R열(총누계), 33행(합계) 같은 수식 셀이 올바른 값을 보여 주려면 필수.
    /// </summary>
    public static void SetFullCalcOnLoad(WorkbookPart workbookPart)
    {
        var calcPr = workbookPart.Workbook
                                 .GetFirstChild<CalculationProperties>();
        if (calcPr == null)
        {
            // CalculationProperties가 없으면 새로 추가.
            // 통합문서 XML 은 요소 순서가 정해져 있어서(… sheets → definedNames → calcPr → … → extLst)
            // 맨 끝에 붙이면 Excel 이 "파일 손상"으로 볼 수 있다. 앞에 와야 하는 요소 중 마지막 것 바로 뒤에 넣는다.
            calcPr = new CalculationProperties();
            var wb = workbookPart.Workbook;
            OpenXmlElement? before = wb.ChildElements.LastOrDefault(e =>
                e is FileVersion or FileSharing or WorkbookProperties or WorkbookProtection or BookViews
                  or Sheets or FunctionGroups or ExternalReferences or DefinedNames);
            if (before != null) wb.InsertAfter(calcPr, before);
            else wb.PrependChild(calcPr);
        }
        calcPr.FullCalculationOnLoad = true;
    }

    // ── WorksheetPart 찾기 ────────────────────────────────────────────

    /// <summary>
    /// 시트 이름으로 WorksheetPart를 찾아 반환한다. 없으면 null.
    /// </summary>
    public static WorksheetPart? FindWorksheetPart(WorkbookPart workbookPart, string sheetName)
    {
        // Workbook의 Sheets 컬렉션에서 이름 일치하는 Sheet를 찾는다
        var sheet = workbookPart.Workbook
                                .Descendants<Sheet>()
                                .FirstOrDefault(s => s.Name?.Value == sheetName);
        if (sheet?.Id?.Value == null) return null;

        return workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
    }

    // ── 내부 헬퍼 (private) ───────────────────────────────────────────

    /// <summary>
    /// 셀 주소로 Cell 엘리먼트를 찾는다. 없으면 null.
    /// </summary>
    private static Cell? FindCell(Worksheet worksheet, string cellAddress)
    {
        (_, uint rowIndex) = ParseAddress(cellAddress);

        // SheetData → Row(행 번호 일치) → Cell(셀 주소 일치)
        return worksheet.GetFirstChild<SheetData>()
                        ?.Elements<Row>()
                         .FirstOrDefault(r => r.RowIndex?.Value == rowIndex)
                        ?.Elements<Cell>()
                         .FirstOrDefault(c => c.CellReference?.Value == cellAddress);
    }

    /// <summary>
    /// 셀을 찾거나 없으면 새로 만든다.
    /// Open XML 규칙: Row 내 Cell은 열 순서(A→B→…)로 삽입해야 한다.
    /// </summary>
    private static Cell GetOrCreateCell(Worksheet worksheet, string cellAddress)
    {
        (string col, uint rowIndex) = ParseAddress(cellAddress);
        var sheetData = worksheet.GetFirstChild<SheetData>()!;

        // ── 행 찾거나 생성 ─────────────────────────
        var row = sheetData.Elements<Row>()
                           .FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = rowIndex };

            // 행도 순서대로 삽입해야 한다 (rowIndex보다 큰 Row 앞에 넣는다)
            var nextRow = sheetData.Elements<Row>()
                                   .FirstOrDefault(r => r.RowIndex?.Value > rowIndex);
            if (nextRow != null)
                sheetData.InsertBefore(row, nextRow);
            else
                sheetData.Append(row);
        }

        // ── 셀 찾거나 생성 ─────────────────────────
        var existingCell = row.Elements<Cell>()
                              .FirstOrDefault(c => c.CellReference?.Value == cellAddress);
        if (existingCell != null) return existingCell;

        // 새 셀: 열 순서를 지켜서 삽입한다
        var newCell = new Cell { CellReference = cellAddress };
        uint newColIndex = ColLetterToIndex(col);

        var refCell = row.Elements<Cell>()
                         .FirstOrDefault(c =>
                         {
                             (string existCol, _) = ParseAddress(c.CellReference?.Value ?? "A1");
                             return ColLetterToIndex(existCol) > newColIndex;
                         });

        if (refCell != null)
            row.InsertBefore(newCell, refCell);
        else
            row.Append(newCell);

        return newCell;
    }

    /// <summary>
    /// SharedStringTable에서 값을 찾거나 없으면 추가하고, 인덱스를 반환한다.
    ///
    /// SharedStringTable: 엑셀 파일 내의 문자열 사전.
    /// "관리자"를 3번 셀에 쓰면 테이블에 한 번만 저장되고, 각 셀은 인덱스 번호만 저장한다.
    /// </summary>
    private static int GetOrAddSharedString(SpreadsheetDocument doc, string value)
    {
        var workbookPart = doc.WorkbookPart!;
        SharedStringTablePart sst;

        if (workbookPart.SharedStringTablePart == null)
            sst = workbookPart.AddNewPart<SharedStringTablePart>();
        else
            sst = workbookPart.SharedStringTablePart;

        if (sst.SharedStringTable == null)
            sst.SharedStringTable = new SharedStringTable();

        var table = sst.SharedStringTable;
        var items = table.Elements<SharedStringItem>().ToList();

        // 이미 같은 값이 있으면 기존 인덱스를 반환
        for (int i = 0; i < items.Count; i++)
        {
            if (items[i].InnerText == value)
                return i;
        }

        // 없으면 새 항목 추가
        table.Append(new SharedStringItem(new Text(value)));
        return items.Count; // 새 인덱스
    }

    /// <summary>
    /// Cell의 실제 텍스트 값을 반환한다.
    /// SharedString이면 테이블에서 조회, 아니면 CellValue 직접 반환.
    /// </summary>
    private static string? ResolveCellValue(SpreadsheetDocument? doc, Cell cell)
    {
        // t="inlineStr": 글자를 칸 안(<is>)에 직접 넣는 방식. Excel 은 잘 안 쓰지만 다른 프로그램(openpyxl 등)이 쓴다
        if (cell.DataType?.Value == CellValues.InlineString)
            return cell.InlineString?.InnerText ?? cell.CellValue?.Text;

        string? raw = cell.CellValue?.Text;
        if (raw == null) return null;

        // t="s" → SharedString 인덱스 → 실제 문자열 조회
        if (cell.DataType?.Value == CellValues.SharedString && doc != null)
        {
            if (int.TryParse(raw, out int idx))
            {
                var sst = doc.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
                if (sst != null)
                {
                    var items = sst.Elements<SharedStringItem>().ToList();
                    if (idx >= 0 && idx < items.Count)
                        return items[idx].InnerText;
                }
            }
        }

        return raw;
    }
}
