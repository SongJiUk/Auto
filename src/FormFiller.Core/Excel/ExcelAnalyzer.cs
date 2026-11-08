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
    /// 지정 시트의 각 셀 배경색을 ARGB (uint) 2D 배열로 반환한다 (0-based index).
    /// 배경색이 없거나 흰색이면 null.
    /// RGB 직접 지정 색상(ForegroundColor.Rgb)만 처리한다.
    /// </summary>
    public static uint?[,] GetCellBackgrounds(string filePath, string sheetName, int maxRow, int maxCol)
    {
        int limitedRows = Math.Min(maxRow, MaxPreviewRows);
        int limitedCols = Math.Min(maxCol, MaxPreviewCols);
        var result = new uint?[limitedRows, limitedCols];

        if (!File.Exists(filePath)) return result;

        try
        {
            using var doc    = SpreadsheetDocument.Open(filePath, isEditable: false);
            var       wbPart = doc.WorkbookPart;
            if (wbPart == null) return result;

            var wsPart = CellHelper.FindWorksheetPart(wbPart, sheetName);
            if (wsPart == null) return result;

            // 스타일 시트에서 CellFormats / Fills 가져오기
            var stylesheet = wbPart.WorkbookStylesPart?.Stylesheet;
            if (stylesheet == null) return result;

            var cellFormats = stylesheet.CellFormats?.Elements<CellFormat>().ToArray() ?? [];
            var fills       = stylesheet.Fills?.Elements<Fill>().ToArray() ?? [];

            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return result;

            foreach (var row in sheetData.Elements<Row>())
            {
                uint rowIdx = row.RowIndex?.Value ?? 0;
                if (rowIdx == 0 || rowIdx > (uint)limitedRows) continue;
                int r = (int)rowIdx - 1;

                foreach (var cell in row.Elements<Cell>())
                {
                    string? cellRef = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(cellRef)) continue;

                    (string colLetter, _) = CellHelper.ParseAddress(cellRef);
                    int colIdx = (int)CellHelper.ColLetterToIndex(colLetter);
                    if (colIdx > limitedCols) continue;
                    int c = colIdx - 1;

                    // 스타일 인덱스 → Fill → PatternFill → FgColor(RGB)
                    if (cell.StyleIndex?.HasValue != true) continue;
                    int styleIdx = (int)cell.StyleIndex.Value;
                    if (styleIdx >= cellFormats.Length) continue;

                    var fmt = cellFormats[styleIdx];
                    if (fmt.FillId?.HasValue != true) continue;

                    int fillIdx = (int)fmt.FillId.Value;
                    if (fillIdx >= fills.Length) continue;

                    var pf = fills[fillIdx].PatternFill;
                    if (pf == null) continue;

                    // 패턴 없음(none)이면 배경색 없음
                    if (pf.PatternType?.Value == PatternValues.None) continue;

                    // FgColor.Rgb 가 있을 때만 처리
                    string? hex = pf.ForegroundColor?.Rgb?.Value;
                    if (string.IsNullOrEmpty(hex)) continue;

                    if (!uint.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out uint argb))
                        continue;

                    // AARRGGBB 형식; AA==00이면 투명(무시)
                    byte alpha = (byte)((argb >> 24) & 0xFF);
                    uint rgb   = argb & 0x00FFFFFF;

                    if (alpha == 0) continue;             // 완전 투명
                    if (rgb == 0xFFFFFF) continue;        // 흰색
                    if (rgb == 0x000000 && alpha < 0x10) continue; // 거의 투명한 검정

                    result[r, c] = argb;
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ExcelAnalyzer] 배경색 읽기 실패: {ex.Message}");
        }

        return result;
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

    // ── 실제 엑셀과 같은 미리보기 데이터 ──────────────────────────────

    /// <summary>
    /// 지정 시트의 미리보기에 필요한 모든 정보를 한 번에 읽어 반환한다.
    ///
    /// 반환 데이터:
    /// - 셀 값, 배경색, 가로 정렬, 볼드, 테두리
    /// - 열 너비(px), 행 높이(px)
    /// - 병합 셀 목록
    ///
    /// 성능 보호를 위해 최대 50행 × 30열로 제한한다.
    /// 오류 발생 시 null 을 반환한다.
    /// </summary>
    public static ExcelPreviewData? GetPreviewData(string filePath, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(filePath) ||
            string.IsNullOrWhiteSpace(sheetName) ||
            !File.Exists(filePath))
            return null;

        try
        {
            using var doc    = SpreadsheetDocument.Open(filePath, isEditable: false);
            var       wbPart = doc.WorkbookPart;
            if (wbPart == null) return null;

            var wsPart = CellHelper.FindWorksheetPart(wbPart, sheetName);
            if (wsPart == null) return null;

            var worksheet = wsPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return null;

            // ── 1. 사용 범위 계산 ──
            (int maxRow, int maxCol) = CalcUsedRange(sheetData);
            int rows = Math.Min(maxRow, MaxPreviewRows);
            int cols = Math.Min(maxCol, MaxPreviewCols);
            if (rows <= 0 || cols <= 0) return null;

            // ── 2. 스타일 배열 로드 ──
            // CellFormats[styleIndex] → FillId, FontId, BorderId, Alignment
            var stylesheet  = wbPart.WorkbookStylesPart?.Stylesheet;
            var cellFormats = stylesheet?.CellFormats?.Elements<CellFormat>().ToArray() ?? [];
            var fills       = stylesheet?.Fills  ?.Elements<Fill>  ().ToArray() ?? [];
            var fonts       = stylesheet?.Fonts  ?.Elements<Font>  ().ToArray() ?? [];
            var borders     = stylesheet?.Borders?.Elements<Border>().ToArray() ?? [];

            // ── 3. 결과 배열 초기화 ──
            var values      = new string[rows, cols];
            var bgColors    = new uint?[rows, cols];
            var hAlignments = new ExcelHAlign[rows, cols];
            var isBold      = new bool[rows, cols];
            var borderInfos = new ExcelBorderInfo[rows, cols];

            // ── 4. 셀 값 + 스타일 읽기 ──
            foreach (var row in sheetData.Elements<Row>())
            {
                uint rowIdx = row.RowIndex?.Value ?? 0;
                if (rowIdx == 0 || rowIdx > (uint)rows) continue;
                int r = (int)rowIdx - 1;

                foreach (var cell in row.Elements<Cell>())
                {
                    string? cellRef = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(cellRef)) continue;

                    (string colLetter, _) = CellHelper.ParseAddress(cellRef);
                    int colIdx = (int)CellHelper.ColLetterToIndex(colLetter);
                    if (colIdx > cols) continue;
                    int c = colIdx - 1;

                    values[r, c] = CellHelper.GetCellValue(wsPart, cellRef) ?? "";

                    if (cell.StyleIndex?.HasValue != true) continue;
                    int styleIdx = (int)cell.StyleIndex.Value;
                    if (styleIdx >= cellFormats.Length) continue;

                    var fmt = cellFormats[styleIdx];

                    // 배경색
                    if (fmt.FillId?.HasValue == true)
                    {
                        int fillIdx = (int)fmt.FillId.Value;
                        if (fillIdx < fills.Length)
                        {
                            var pf = fills[fillIdx].PatternFill;
                            if (pf != null && pf.PatternType?.Value != PatternValues.None)
                            {
                                string? hex = pf.ForegroundColor?.Rgb?.Value;
                                if (!string.IsNullOrEmpty(hex) &&
                                    uint.TryParse(hex, System.Globalization.NumberStyles.HexNumber,
                                                  null, out uint argb))
                                {
                                    byte alpha = (byte)((argb >> 24) & 0xFF);
                                    uint rgb   = argb & 0x00FFFFFF;
                                    if (alpha != 0 && rgb != 0xFFFFFF)
                                        bgColors[r, c] = argb;
                                }
                            }
                        }
                    }

                    // 가로 정렬
                    if (fmt.Alignment?.Horizontal?.HasValue == true)
                    {
                        var h = fmt.Alignment.Horizontal.Value;
                        if      (h == HorizontalAlignmentValues.Left)              hAlignments[r, c] = ExcelHAlign.Left;
                        else if (h == HorizontalAlignmentValues.Center)            hAlignments[r, c] = ExcelHAlign.Center;
                        else if (h == HorizontalAlignmentValues.CenterContinuous)  hAlignments[r, c] = ExcelHAlign.Center;
                        else if (h == HorizontalAlignmentValues.Right)             hAlignments[r, c] = ExcelHAlign.Right;
                        else if (h == HorizontalAlignmentValues.Fill)              hAlignments[r, c] = ExcelHAlign.Fill;
                        else if (h == HorizontalAlignmentValues.Justify)           hAlignments[r, c] = ExcelHAlign.Justify;
                    }

                    // 볼드
                    if (fmt.FontId?.HasValue == true)
                    {
                        int fontIdx = (int)fmt.FontId.Value;
                        if (fontIdx < fonts.Length)
                        {
                            var font = fonts[fontIdx];
                            isBold[r, c] = font.Bold != null && (font.Bold.Val?.Value ?? true);
                        }
                    }

                    // 테두리
                    if (fmt.BorderId?.HasValue == true)
                    {
                        int borderIdx = (int)fmt.BorderId.Value;
                        if (borderIdx < borders.Length)
                        {
                            var bd = borders[borderIdx];
                            borderInfos[r, c] = new ExcelBorderInfo
                            {
                                Left   = HasBorderStyle(bd.LeftBorder),
                                Top    = HasBorderStyle(bd.TopBorder),
                                Right  = HasBorderStyle(bd.RightBorder),
                                Bottom = HasBorderStyle(bd.BottomBorder),
                            };
                        }
                    }
                }
            }

            // ── 5. 열 너비 (Excel width × 7px) ──
            var sfp              = worksheet.SheetFormatProperties;
            double defColWidth   = sfp?.DefaultColumnWidth?.Value ?? 8.0;
            var colWidths        = new double[cols];
            for (int c = 0; c < cols; c++)
                colWidths[c] = Math.Clamp(defColWidth * 7.0, 20, 200);

            var columnsElem = worksheet.GetFirstChild<Columns>();
            if (columnsElem != null)
            {
                foreach (var col in columnsElem.Elements<Column>())
                {
                    int    minC  = (int)(col.Min?.Value ?? 0);
                    int    maxC  = (int)(col.Max?.Value ?? 0);
                    double width = col.Width?.Value ?? defColWidth;
                    double px    = Math.Clamp(width * 7.0, 20, 200);

                    if (minC <= 0) continue; // Min 없는 열은 건너뜀

                    for (int ci = minC; ci <= maxC && ci <= cols; ci++)
                        colWidths[ci - 1] = px;
                }
            }

            // ── 6. 행 높이 (Excel points × 4/3 → px) ──
            double defRowHeight = sfp?.DefaultRowHeight?.Value ?? 15.0;
            var rowHeights      = new double[rows];
            for (int r = 0; r < rows; r++)
                rowHeights[r] = Math.Clamp(defRowHeight * 4.0 / 3.0, 16, 100);

            foreach (var row in sheetData.Elements<Row>())
            {
                uint rowIdx = row.RowIndex?.Value ?? 0;
                if (rowIdx == 0 || rowIdx > (uint)rows) continue;
                int r = (int)rowIdx - 1;

                if (row.CustomHeight?.Value == true && row.Height?.HasValue == true)
                    rowHeights[r] = Math.Clamp(row.Height.Value * 4.0 / 3.0, 16, 100);
            }

            // ── 7. 병합 셀 ──
            var merges        = new List<ExcelMergeRange>();
            var mergeCellsElem = worksheet.GetFirstChild<MergeCells>();
            if (mergeCellsElem != null)
            {
                foreach (var mc in mergeCellsElem.Elements<MergeCell>())
                {
                    var range = ParseMergeRange(mc.Reference?.Value);
                    if (range != null &&
                        range.TopRow  <= rows &&
                        range.LeftCol <= cols)
                        merges.Add(range);
                }
            }

            return new ExcelPreviewData
            {
                Rows         = rows,
                Cols         = cols,
                Values       = values,
                BgColors     = bgColors,
                HAlignments  = hAlignments,
                IsBold       = isBold,
                Borders      = borderInfos,
                ColWidthsPx  = colWidths,
                RowHeightsPx = rowHeights,
                Merges       = merges,
            };
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[ExcelAnalyzer] 미리보기 데이터 읽기 실패: {ex.Message}");
            return null;
        }
    }

    // ── 내부 헬퍼 ──────────────────────────────────────────────────────

    /// <summary>SheetData 를 순회해 실제 사용 범위(최대 행·열 1-based)를 계산한다.</summary>
    private static (int maxRow, int maxCol) CalcUsedRange(SheetData sheetData)
    {
        int maxRow = 0, maxCol = 0;
        foreach (var row in sheetData.Elements<Row>())
        {
            uint rowIdx = row.RowIndex?.Value ?? 0;
            if (rowIdx > maxRow) maxRow = (int)rowIdx;

            foreach (var cell in row.Elements<Cell>())
            {
                string? cellRef = cell.CellReference?.Value;
                if (string.IsNullOrEmpty(cellRef)) continue;
                (string colLetter, _) = CellHelper.ParseAddress(cellRef);
                int colIdx = (int)CellHelper.ColLetterToIndex(colLetter);
                if (colIdx > maxCol) maxCol = colIdx;
            }
        }
        return (maxRow, maxCol);
    }

    /// <summary>
    /// BorderPropertiesType (Left/Right/Top/BottomBorder) 의 Style 이 None 이 아닌지 확인한다.
    /// Style 이 없거나 None 이면 false 를 반환한다.
    /// </summary>
    private static bool HasBorderStyle(BorderPropertiesType? bp) =>
        bp?.Style?.HasValue == true && bp.Style.Value != BorderStyleValues.None;

    /// <summary>
    /// "A1:B2" 형식의 병합 셀 범위 문자열을 ExcelMergeRange 로 변환한다.
    /// 형식이 올바르지 않으면 null 을 반환한다.
    /// </summary>
    private static ExcelMergeRange? ParseMergeRange(string? reference)
    {
        if (string.IsNullOrEmpty(reference)) return null;

        var parts = reference.Split(':');
        if (parts.Length != 2) return null;

        (string col1Str, uint row1) = CellHelper.ParseAddress(parts[0]);
        (string col2Str, uint row2) = CellHelper.ParseAddress(parts[1]);

        if (row1 == 0 || row2 == 0) return null;

        int r1 = (int)row1;
        int r2 = (int)row2;
        int c1 = (int)CellHelper.ColLetterToIndex(col1Str);
        int c2 = (int)CellHelper.ColLetterToIndex(col2Str);
        if (c1 <= 0 || c2 <= 0) return null;

        return new ExcelMergeRange
        {
            TopRow    = Math.Min(r1, r2),
            BottomRow = Math.Max(r1, r2),
            LeftCol   = Math.Min(c1, c2),
            RightCol  = Math.Max(c1, c2),
        };
    }
}
