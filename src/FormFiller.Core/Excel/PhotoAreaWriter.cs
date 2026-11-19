using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FormFiller.Core.Config;
using FormFiller.Core.Models;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace FormFiller.Core.Excel;

/// <summary>
/// 사진 구역 블록(BlockTypes.Photos) 쓰기.
///
/// [하는 일]
/// 1. 그날 넣을 사진을 고른다 (파일이 있는 것만, 칸 수까지)
/// 2. 넣을 자리를 정한다: 1장이고 "1장일 때 자리"가 있으면 그 자리, 아니면 칸을 순서대로
/// 3. 구역(모든 칸 + 1장 자리) 안에서 시작하는 그림을 지운다 → 정한 자리에 새 그림을 넣는다
/// 4. 자리마다 설명 칸에 설명을 쓴다
///
/// [기존 사진 교체(ReplacePictureAt)와 다른 점]
/// 교체는 서식에 있는 그림의 이미지만 바꾸므로 그림 개수·위치가 서식에 묶인다.
/// 여기서는 그림 자체를 지우고 새로 만들어서, 서식이 1장 모양이든 2장 모양이든 결과가 같다.
/// </summary>
public static class PhotoAreaWriter
{
    private const long EmuPerPixel = 9525;          // 914400 EMU = 1인치 = 96px
    private const double BaseColumnWidth = 8;       // 엑셀 기본 열 너비 (여백 뺀 글자 수)
    private const int ColumnPaddingPx = 5;          // 열 양쪽 여백 + 눈금선 (px)
    private const int DefaultWidthStepPx = 8;       // 기본 열 너비는 8px 단위로 올림
    private const double DefaultRowHeightPt = 15.0; // 엑셀 기본 행 높이 (포인트)
    private const double DefaultFontSizePt = 11;

    /// <summary>
    /// 글꼴별 숫자 한 글자 너비 (11pt 기준 px). 열 너비(글자 수)를 px 로 바꿀 때 쓴다.
    /// 맑은 고딕 = 한국어 엑셀 기본 (기본 열 72px), Calibri = 영어 엑셀 기본 (기본 열 64px).
    /// 목록에 없는 글꼴은 Calibri 값으로 계산한다 (사진 비율이 조금 다를 수 있음).
    /// </summary>
    private static readonly Dictionary<string, double> DigitWidthAt11Pt = new(StringComparer.OrdinalIgnoreCase)
    {
        ["맑은 고딕"] = 8, ["Malgun Gothic"] = 8, ["Calibri"] = 7,
    };
    private const double FallbackDigitWidthAt11Pt = 7;

    public static void Write(FieldConfig f, SpreadsheetDocument pkg, WorksheetPart ws, BlockValue v, string sheet,
                             Func<string, int, int, byte[]>? photoLoader, List<string>? notes)
    {
        var areas = (f.PhotoAreas ?? []).Where(a => !string.IsNullOrWhiteSpace(a.Range)).ToList();
        if (areas.Count == 0) return;

        var photos = new List<PhotoValue>();
        foreach (var p in v.Photos.Where(p => !string.IsNullOrWhiteSpace(p.Path)))
        {
            if (File.Exists(p.Path)) photos.Add(p);
            else notes?.Add($"[{sheet}] {f.Label}: 사진 파일이 없어 건너뜀 ({Path.GetFileName(p.Path)})");
        }

        if (photos.Count == 0 || photoLoader == null)
        {
            notes?.Add($"[{sheet}] {f.Label}: 사진 없음 → 바탕 파일 사진 그대로");
            return;
        }
        if (photos.Count > areas.Count)
        {
            notes?.Add($"[{sheet}] {f.Label}: 사진 {photos.Count}장 중 칸이 {areas.Count}칸이라 앞 {areas.Count}장만 넣음");
            photos = photos.Take(areas.Count).ToList();
        }

        var targets = photos.Count == 1 && !string.IsNullOrWhiteSpace(f.SinglePhotoArea?.Range)
            ? [f.SinglePhotoArea!]
            : areas.Take(photos.Count).ToList();

        var drawing = EnsureDrawingsPart(ws);
        RemovePicturesIn(drawing, AllRanges(f));

        for (int i = 0; i < targets.Count; i++)
        {
            AddPicture(ws, drawing, targets[i].Range, photos[i].Path, photoLoader);
            if (!string.IsNullOrWhiteSpace(targets[i].CaptionCell) && !CellHelper.HasFormula(ws, targets[i].CaptionCell!))
                CellHelper.SetCellString(pkg, ws, targets[i].CaptionCell!, photos[i].Caption);
        }
        drawing.WorksheetDrawing.Save();
    }

    /// <summary>앞 날짜 파일(없으면 서식)의 설명 칸 글자를 칸 순서대로 읽는다. 사진 경로는 비워 둔다(매일 새로 고름).</summary>
    public static List<PhotoValue> ReadCaptions(FieldConfig f, WorksheetPart ws) =>
        (f.PhotoAreas ?? []).Select(a => new PhotoValue
        {
            Caption = string.IsNullOrWhiteSpace(a.CaptionCell) ? "" : CellHelper.GetCellValue(ws, a.CaptionCell!) ?? "",
        }).ToList();

    /// <summary>블록의 모든 자리 범위 (칸들 + 1장 자리)</summary>
    public static List<string> AllRanges(FieldConfig f) =>
        (f.PhotoAreas ?? []).Append(f.SinglePhotoArea).OfType<PhotoArea>()
            .Select(a => a.Range).Where(r => !string.IsNullOrWhiteSpace(r)).ToList();

    /// <summary>그림 앵커의 왼쪽 위가 범위들 중 하나 안에 있는지 (검증에서 "바뀌어도 되는 그림"을 가릴 때도 쓴다)</summary>
    public static bool StartsIn(OpenXmlCompositeElement anchor, IEnumerable<string> ranges)
    {
        if (anchor.GetFirstChild<Xdr.FromMarker>() is not { } from) return false;
        int col = int.TryParse(from.ColumnId?.Text, out var c) ? c : -1;
        int row = int.TryParse(from.RowId?.Text, out var r) ? r : -1;
        return ranges.Select(ParseRange).Any(g => col >= g.C1 && col <= g.C2 && row >= g.R1 && row <= g.R2);
    }

    // ── 그림 지우기·넣기 ─────────────────────────────────────────────

    private static void RemovePicturesIn(DrawingsPart drawing, List<string> ranges)
    {
        var doomed = drawing.WorksheetDrawing.ChildElements.OfType<OpenXmlCompositeElement>()
            .Where(a => a is Xdr.TwoCellAnchor or Xdr.OneCellAnchor or Xdr.AbsoluteAnchor)
            .Where(a => a.Descendants<Xdr.Picture>().Any() && StartsIn(a, ranges))
            .ToList();

        var oldIds = doomed.SelectMany(a => a.Descendants<A.Blip>()).Select(b => b.Embed?.Value)
                           .OfType<string>().Distinct().ToList();
        foreach (var a in doomed) a.Remove();

        // 다른 그림이 같이 쓰지 않는 이미지만 지운다 (앞 날짜 파일을 바탕으로 만들 때 파일이 커지지 않게)
        var stillUsed = drawing.WorksheetDrawing.Descendants<A.Blip>().Select(b => b.Embed?.Value).ToHashSet();
        foreach (string id in oldIds.Where(id => !stillUsed.Contains(id)))
        {
            try { drawing.DeletePart(id); } catch { /* 못 지우면 그대로 둔다 */ }
        }
    }

    private static void AddPicture(WorksheetPart ws, DrawingsPart drawing, string range, string photoPath,
                                   Func<string, int, int, byte[]> photoLoader)
    {
        var g = ParseRange(range);
        double digit = MaxDigitWidthPx(ws);
        long cx = 0, cy = 0;
        for (int c = g.C1; c <= g.C2; c++) cx += ColumnWidthPx(ws, c, digit) * EmuPerPixel;
        for (int r = g.R1; r <= g.R2; r++) cy += RowHeightPx(ws, r) * EmuPerPixel;

        // 인쇄 품질을 위해 화면 크기(96 DPI)의 2배 해상도로 만든다 (ReplacePictureAt 과 같은 기준)
        byte[] bytes = photoLoader(photoPath, (int)(cx / EmuPerPixel * 2), (int)(cy / EmuPerPixel * 2));
        var image = drawing.AddImagePart(ImagePartType.Jpeg);
        using (var ms = new MemoryStream(bytes)) image.FeedData(ms);

        uint id = NextShapeId(drawing);
        var anchor = new Xdr.TwoCellAnchor(
            Marker<Xdr.FromMarker>(g.C1, g.R1),
            Marker<Xdr.ToMarker>(g.C2 + 1, g.R2 + 1),   // 범위 오른쪽 아래 끝 = 다음 칸의 시작
            new Xdr.Picture(
                new Xdr.NonVisualPictureProperties(
                    new Xdr.NonVisualDrawingProperties { Id = id, Name = $"사진 {id}" },
                    new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
                new Xdr.BlipFill(
                    new A.Blip { Embed = drawing.GetIdOfPart(image) },
                    new A.Stretch(new A.FillRectangle())),
                new Xdr.ShapeProperties(
                    new A.Transform2D(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = cx, Cy = cy }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
            new Xdr.ClientData())
        { EditAs = Xdr.EditAsValues.OneCell };
        drawing.WorksheetDrawing.Append(anchor);
    }

    private static T Marker<T>(int col, int row) where T : OpenXmlCompositeElement, new()
    {
        var m = new T();
        m.Append(new Xdr.ColumnId(col.ToString()), new Xdr.ColumnOffset("0"),
                 new Xdr.RowId(row.ToString()), new Xdr.RowOffset("0"));
        return m;
    }

    private static uint NextShapeId(DrawingsPart drawing) =>
        drawing.WorksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>()
               .Select(p => p.Id?.Value ?? 0).DefaultIfEmpty(1u).Max() + 1;

    /// <summary>시트에 그림 묶음(drawing)이 없으면 만든다. 시트 XML 안 &lt;drawing&gt; 위치는 스키마 순서를 지킨다.</summary>
    private static DrawingsPart EnsureDrawingsPart(WorksheetPart ws)
    {
        if (ws.DrawingsPart is { } existing) return existing;

        var dp = ws.AddNewPart<DrawingsPart>();
        dp.WorksheetDrawing = new Xdr.WorksheetDrawing();
        dp.WorksheetDrawing.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
        dp.WorksheetDrawing.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

        var element = new Drawing { Id = ws.GetIdOfPart(dp) };
        var after = ws.Worksheet.ChildElements.FirstOrDefault(e =>
            e is LegacyDrawing or LegacyDrawingHeaderFooter or DrawingHeaderFooter or Picture or OleObjects
              or Controls or WebPublishItems or TableParts or WorksheetExtensionList);
        if (after != null) ws.Worksheet.InsertBefore(element, after);
        else ws.Worksheet.Append(element);
        return dp;
    }

    // ── 칸 크기 ──────────────────────────────────────────────────────

    /// <summary>
    /// 열 너비 (px). 파일에 저장된 너비(col width, defaultColWidth)는 여백이 이미 들어간 글자 수라
    /// ECMA-376 공식 Truncate(((256 × 너비 + Truncate(128 / 숫자너비)) / 256) × 숫자너비) 를 쓴다.
    /// 둘 다 없으면 기본 너비 = 8글자 × 숫자너비 + 여백을 8px 단위로 올림.
    /// </summary>
    private static long ColumnWidthPx(WorksheetPart ws, int col0, double digit)
    {
        uint col1 = (uint)col0 + 1;
        double? width = ws.Worksheet.GetFirstChild<Columns>()?.Elements<Column>()
                            .FirstOrDefault(c => c.Min?.Value <= col1 && c.Max?.Value >= col1)?.Width?.Value
                        ?? ws.Worksheet.GetFirstChild<SheetFormatProperties>()?.DefaultColumnWidth?.Value;
        if (width is { } w)
            return (long)Math.Truncate((256 * w + Math.Truncate(128 / digit)) / 256 * digit);

        double px = BaseColumnWidth * digit + ColumnPaddingPx;
        return (long)(Math.Ceiling(px / DefaultWidthStepPx) * DefaultWidthStepPx);
    }

    /// <summary>통합 문서 기본 글꼴(스타일의 첫 글꼴)의 숫자 한 글자 너비 (px)</summary>
    private static double MaxDigitWidthPx(WorksheetPart ws)
    {
        var wb = ws.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
        var font = wb?.WorkbookStylesPart?.Stylesheet?.Fonts?.Elements<Font>().FirstOrDefault();
        string name = font?.FontName?.Val?.Value ?? "";
        double size = font?.FontSize?.Val?.Value ?? DefaultFontSizePt;
        double at11 = DigitWidthAt11Pt.TryGetValue(name, out var d) ? d : FallbackDigitWidthAt11Pt;
        return Math.Max(1, Math.Round(at11 * size / DefaultFontSizePt));
    }

    /// <summary>행 높이 (px). 포인트 × 4/3</summary>
    private static long RowHeightPx(WorksheetPart ws, int row0)
    {
        uint row1 = (uint)row0 + 1;
        var row = ws.Worksheet.GetFirstChild<SheetData>()?.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == row1);
        double pt = row?.CustomHeight?.Value == true && row.Height?.HasValue == true ? row.Height.Value
                  : ws.Worksheet.GetFirstChild<SheetFormatProperties>()?.DefaultRowHeight?.Value ?? DefaultRowHeightPt;
        return (long)Math.Round(pt * 4 / 3);
    }

    /// <summary>"A4:H15" → 0부터 세는 (열1, 행1, 열2, 행2). 칸 하나("B16")도 된다.</summary>
    private static (int C1, int R1, int C2, int R2) ParseRange(string range)
    {
        var parts = range.Split(':', StringSplitOptions.TrimEntries);
        var (c1, r1) = Cell0(parts[0]);
        var (c2, r2) = parts.Length > 1 ? Cell0(parts[1]) : (c1, r1);
        return (Math.Min(c1, c2), Math.Min(r1, r2), Math.Max(c1, c2), Math.Max(r1, r2));
    }

    private static (int Col, int Row) Cell0(string cell)
    {
        var (col, row) = CellHelper.ParseAddress(cell.Replace("$", ""));
        return ((int)CellHelper.ColLetterToIndex(col) - 1, (int)row - 1);
    }
}
