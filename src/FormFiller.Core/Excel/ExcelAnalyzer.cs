using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FormFiller.Core.Excel;

/// <summary>
/// Excel 파일에서 시트 목록이나 셀 미리보기 값을 읽어 오는 분석 도우미.
/// 설정 화면에서 "어떤 시트가 있지?", "이 셀에 뭐가 들어 있지?" 를 확인할 때 사용한다.
///
/// [Open XML 읽기 전용 열기 패턴]
/// SpreadsheetDocument.Open(path, isEditable: false) 로 열면
/// 파일 잠금 없이 읽기만 가능하다 → 사용자가 파일을 Excel 에서 열어 뒀어도 읽힌다.
/// </summary>
public static class ExcelAnalyzer
{
    // ── 미리보기용 최대 범위 제한 (성능) ──────────────────────────────
    private const int MaxPreviewRows = 50;
    private const int MaxPreviewCols = 30;
    /// <summary>
    /// 지정 Excel 파일에 들어 있는 모든 시트 이름을 순서대로 반환한다.
    /// 파일이 없거나 오류가 나면 빈 목록을 반환한다.
    /// </summary>
    /// <param name="filePath">읽을 .xlsx 파일 전체 경로</param>
    public static List<string> GetSheetNames(string filePath)
    {
        if (!File.Exists(filePath))
            return [];

        try
        {
            // isEditable: false → 읽기 전용으로 열어 원본 훼손 방지
            using var doc = SpreadsheetDocument.Open(filePath, isEditable: false);

            // Workbook > Sheets > Sheet 각각의 Name 속성을 꺼낸다
            return doc.WorkbookPart?.Workbook
                      .Descendants<Sheet>()
                      .Select(s => s.Name?.Value ?? "")
                      .Where(name => name.Length > 0)
                      .ToList()
                   ?? [];
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ExcelAnalyzer] 시트 목록 읽기 실패: {ex.Message}");
            return [];
        }
    }

    /// <summary>
    /// 특정 시트의 특정 셀 값을 문자열로 반환한다 (설정 화면 미리보기용).
    ///
    /// 반환 규칙:
    /// - 셀이 비어 있으면 ""
    /// - filePath 나 sheetName 이 비어 있으면 ""
    /// - 오류 발생 시 "(읽기 오류)"
    /// </summary>
    /// <param name="filePath">읽을 .xlsx 파일 전체 경로</param>
    /// <param name="sheetName">읽을 시트 이름</param>
    /// <param name="cellAddress">읽을 셀 주소 (예: "C6", "AD6")</param>
    public static string GetCellPreview(string filePath, string sheetName, string cellAddress)
    {
        // 필수값 미입력이면 빈 문자열 (에러 없이)
        if (string.IsNullOrWhiteSpace(filePath)   ||
            string.IsNullOrWhiteSpace(sheetName)  ||
            string.IsNullOrWhiteSpace(cellAddress))
            return "";

        if (!File.Exists(filePath))
            return "";

        try
        {
            using var doc = SpreadsheetDocument.Open(filePath, isEditable: false);

            var workbookPart = doc.WorkbookPart;
            if (workbookPart == null) return "";

            // CellHelper 의 FindWorksheetPart 를 재활용해 WorksheetPart 를 찾는다
            var wsPart = CellHelper.FindWorksheetPart(workbookPart, sheetName);
            if (wsPart == null) return "";

            // CellHelper.GetCellValue 로 실제 텍스트를 가져온다
            return CellHelper.GetCellValue(wsPart, cellAddress) ?? "";
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[ExcelAnalyzer] 셀 미리보기 실패 ({cellAddress}): {ex.Message}");
            return "(읽기 오류)";
        }
    }

    /// <summary>
    /// 지정 시트의 실제 사용 범위(행·열 최댓값)를 반환한다.
    /// SheetData 의 모든 Row 를 순회하며 가장 큰 RowIndex 와 열 인덱스를 찾는다.
    ///
    /// 반환값은 1-based (예: maxRow=36, maxCol=44 이면 36행 · AV열까지 데이터 있음).
    /// 파일이 없거나 오류 시 (0, 0) 을 반환한다.
    /// </summary>
    /// <param name="filePath">읽을 .xlsx 파일 전체 경로</param>
    /// <param name="sheetName">읽을 시트 이름</param>
    public static (int maxRow, int maxCol) GetUsedRange(string filePath, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(filePath) ||
            string.IsNullOrWhiteSpace(sheetName) ||
            !File.Exists(filePath))
            return (0, 0);

        try
        {
            using var doc = SpreadsheetDocument.Open(filePath, isEditable: false);
            var workbookPart = doc.WorkbookPart;
            if (workbookPart == null) return (0, 0);

            var wsPart = CellHelper.FindWorksheetPart(workbookPart, sheetName);
            if (wsPart == null) return (0, 0);

            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return (0, 0);

            int maxRow = 0;
            int maxCol = 0;

            // 모든 Row 를 순회해 최대 RowIndex 와 최대 열 인덱스를 찾는다
            foreach (var row in sheetData.Elements<Row>())
            {
                uint rowIdx = row.RowIndex?.Value ?? 0;
                if (rowIdx > maxRow)
                    maxRow = (int)rowIdx;

                // 각 셀의 열 문자를 열 인덱스로 변환해 최댓값 갱신
                foreach (var cell in row.Elements<Cell>())
                {
                    string? cellRef = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(cellRef)) continue;

                    (string colLetter, _) = CellHelper.ParseAddress(cellRef);
                    int colIdx = (int)CellHelper.ColLetterToIndex(colLetter);
                    if (colIdx > maxCol)
                        maxCol = colIdx;
                }
            }

            return (maxRow, maxCol);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[ExcelAnalyzer] 사용 범위 읽기 실패: {ex.Message}");
            return (0, 0);
        }
    }

    /// <summary>
    /// 지정 시트의 셀 값을 2D 배열 [row, col] 로 반환한다 (0-based index).
    /// row 0 = 1행, col 0 = A열에 해당한다.
    ///
    /// maxRow / maxCol 은 GetUsedRange 로 얻은 값을 그대로 전달한다.
    /// 성능 보호를 위해 최대 50행 × 30열로 제한한다.
    /// CellHelper.GetCellValue 를 활용해 SharedString 도 올바르게 처리한다.
    /// </summary>
    /// <param name="filePath">읽을 .xlsx 파일 전체 경로</param>
    /// <param name="sheetName">읽을 시트 이름</param>
    /// <param name="maxRow">읽을 최대 행 번호 (1-based)</param>
    /// <param name="maxCol">읽을 최대 열 번호 (1-based)</param>
    public static string[,] GetCellValues(string filePath, string sheetName, int maxRow, int maxCol)
    {
        // 최대 범위로 제한
        int limitedRows = Math.Min(maxRow, MaxPreviewRows);
        int limitedCols = Math.Min(maxCol, MaxPreviewCols);

        // 범위가 0 이하이면 빈 배열 반환
        if (limitedRows <= 0 || limitedCols <= 0)
            return new string[0, 0];

        // 결과 배열 초기화 (0-based: [0..limitedRows-1, 0..limitedCols-1])
        var result = new string[limitedRows, limitedCols];

        if (string.IsNullOrWhiteSpace(filePath) ||
            string.IsNullOrWhiteSpace(sheetName) ||
            !File.Exists(filePath))
            return result;

        try
        {
            using var doc = SpreadsheetDocument.Open(filePath, isEditable: false);
            var workbookPart = doc.WorkbookPart;
            if (workbookPart == null) return result;

            var wsPart = CellHelper.FindWorksheetPart(workbookPart, sheetName);
            if (wsPart == null) return result;

            // 범위 내 각 행을 순회하며 셀 값을 배열에 채운다
            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return result;

            foreach (var row in sheetData.Elements<Row>())
            {
                uint rowIdx = row.RowIndex?.Value ?? 0;
                if (rowIdx == 0 || rowIdx > (uint)limitedRows) continue;

                int r = (int)rowIdx - 1; // 0-based row 인덱스

                foreach (var cell in row.Elements<Cell>())
                {
                    string? cellRef = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(cellRef)) continue;

                    (string colLetter, _) = CellHelper.ParseAddress(cellRef);
                    int colIdx = (int)CellHelper.ColLetterToIndex(colLetter);
                    if (colIdx > limitedCols) continue;

                    int c = colIdx - 1; // 0-based col 인덱스

                    // CellHelper.GetCellValue 로 SharedString 포함 모든 타입 처리
                    string? val = CellHelper.GetCellValue(wsPart, cellRef);
                    result[r, c] = val ?? "";
                }
            }

            return result;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[ExcelAnalyzer] 셀 값 읽기 실패: {ex.Message}");
            return result;
        }
    }

    /// <summary>
    /// 1-based 열 인덱스를 열 문자열로 변환한다 (예: 1→"A", 26→"Z", 27→"AA").
    /// GetCellValues 에서 미리보기 열 헤더를 생성할 때 사용한다.
    /// </summary>
    public static string ColIndexToLetter(int colIndex)
    {
        string result = "";
        while (colIndex > 0)
        {
            int rem = (colIndex - 1) % 26;
            result = (char)('A' + rem) + result;
            colIndex = (colIndex - 1) / 26;
        }
        return result;
    }
}
