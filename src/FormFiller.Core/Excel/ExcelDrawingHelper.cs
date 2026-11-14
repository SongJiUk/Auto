using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FormFiller.Core.Config;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace FormFiller.Core.Excel;

/// <summary>
/// 시트 위의 그림·글상자(drawing)를 다룬다: 칸 위치로 그림 찾기, 사진 교체, 그림 옮기기.
///
/// [엑셀 그림 위치의 원리]
/// xl/drawings/drawingN.xml 안에 그림마다 "앵커"가 있다.
///   from = 그림 왼쪽 위가 몇 번째 열·행(0부터)의 칸에서, 칸 안으로 얼마나(EMU) 떨어져 있는지
///   to   = 그림 오른쪽 아래 (twoCellAnchor 만 있음. oneCellAnchor 는 from + 크기(ext))
/// 그래서 그림을 옮기려면 from/to 숫자만 바꾸면 되고, 그림 자체는 건드리지 않는다.
/// </summary>
public static class ExcelDrawingHelper
{
    /// <summary>EMU → 96 DPI 픽셀 (914400 EMU = 1인치 = 96px)</summary>
    public const double EmuPerPixel = 914400.0 / 96.0;

    // ── 찾기 ─────────────────────────────────────────────────────────

    /// <summary>시트의 모든 앵커(twoCell / oneCell / absolute)</summary>
    private static IEnumerable<OpenXmlCompositeElement> Anchors(WorksheetPart wsPart) =>
        wsPart.DrawingsPart?.WorksheetDrawing?.ChildElements.OfType<OpenXmlCompositeElement>()
            .Where(e => e is Xdr.TwoCellAnchor or Xdr.OneCellAnchor or Xdr.AbsoluteAnchor)
        ?? [];

    private static Xdr.FromMarker? From(OpenXmlCompositeElement a) => a.GetFirstChild<Xdr.FromMarker>();
    private static Xdr.ToMarker?   To(OpenXmlCompositeElement a)   => a.GetFirstChild<Xdr.ToMarker>();

    private static int  Int(OpenXmlElement? e)  => int.TryParse(e?.InnerText, out var v) ? v : 0;
    private static long Long(OpenXmlElement? e) => long.TryParse(e?.InnerText, out var v) ? v : 0;

    /// <summary>앵커 안 첫 번째 그림·도형의 이름 (예: "그림 7", "TextBox 10")</summary>
    public static string? ShapeName(OpenXmlCompositeElement anchor) =>
        anchor.Descendants<Xdr.NonVisualDrawingProperties>().FirstOrDefault()?.Name?.Value;

    /// <summary>
    /// 그림 이름이 같은지. 한국어 Excel 은 파일에 "그림 21" 로 저장하지만,
    /// Excel 화면(COM)에서는 같은 그림을 "Picture 21" 처럼 영어 기본 이름으로 알려 준다.
    /// 기본 이름 뒤 번호는 시트 안에서 하나뿐이므로, 이름이 다르면 끝 번호로 비교한다.
    /// </summary>
    public static bool NameMatches(string? fileName, string wanted)
    {
        if (fileName == null) return false;
        if (string.Equals(fileName, wanted, StringComparison.OrdinalIgnoreCase)) return true;
        string? a = TrailingNumber(fileName), b = TrailingNumber(wanted);
        return a != null && a == b;
    }

    private static string? TrailingNumber(string name)
    {
        int i = name.LastIndexOf(' ');
        return i > 0 && i < name.Length - 1 && name[(i + 1)..].All(char.IsDigit) ? name[(i + 1)..] : null;
    }

    /// <summary>이름으로 앵커 찾기: 정확히 같은 이름을 먼저, 없으면 끝 번호가 같은 것</summary>
    private static OpenXmlCompositeElement? FindByName(List<OpenXmlCompositeElement> anchors, string name) =>
        anchors.FirstOrDefault(a => string.Equals(ShapeName(a), name, StringComparison.OrdinalIgnoreCase))
        ?? anchors.FirstOrDefault(a => NameMatches(ShapeName(a), name));

    /// <summary>
    /// 왼쪽 위가 cellAddress 칸에서 시작하는 그림(Picture)의 앵커를 찾는다.
    /// 사진 칸은 시트마다 그림 이름이 달라도 놓인 칸은 같으므로, 칸 위치로 찾는다.
    /// </summary>
    public static OpenXmlCompositeElement? FindPictureAt(WorksheetPart wsPart, string cellAddress)
    {
        var (col, row) = CellHelper.ParseAddress(cellAddress);
        int c = (int)CellHelper.ColLetterToIndex(col) - 1, r = (int)row - 1;

        var pics = Anchors(wsPart).Where(a => a.Descendants<Xdr.Picture>().Any() && From(a) != null).ToList();
        return pics.FirstOrDefault(a => Int(From(a)!.ColumnId) == c && Int(From(a)!.RowId) == r)
            // 손으로 붙인 사진은 한 칸 위·아래에서 시작하기도 한다 (예: 여성회관 5번 사진이 D24 가 아니라 D23).
            // 같은 열에서 한 줄 차이까지는 같은 칸으로 본다.
            ?? pics.Where(a => Int(From(a)!.ColumnId) == c && Math.Abs(Int(From(a)!.RowId) - r) <= 1)
                   .OrderBy(a => Math.Abs(Int(From(a)!.RowId) - r))
                   .FirstOrDefault();
    }

    /// <summary>그림 크기(EMU). twoCellAnchor 도 그림 속성(xfrm)의 크기를 쓴다.</summary>
    public static (long cx, long cy) GetPictureSize(OpenXmlCompositeElement anchor)
    {
        var ext = anchor.Descendants<Xdr.Picture>().FirstOrDefault()
                        ?.ShapeProperties?.Transform2D?.Extents;
        return (ext?.Cx?.Value ?? 0, ext?.Cy?.Value ?? 0);
    }

    // ── 사진 교체 ────────────────────────────────────────────────────

    /// <summary>
    /// cellAddress 에 놓인 그림을 새 이미지로 바꾼다. 위치·크기·테두리는 그대로 둔다.
    ///
    /// [기존 ExcelPhotoWriter.ReplacePhoto 와 다른 점]
    /// 엑셀은 같은 이미지를 여러 그림이 "공유"하도록 저장하기도 한다(빈 칸 표시 이미지 등).
    /// 공유된 이미지 데이터를 덮어쓰면 다른 칸 사진까지 바뀌므로,
    /// 여기서는 새 이미지 파트를 만들어 이 그림만 새 이미지를 가리키게 한다.
    /// </summary>
    /// <param name="loadImage">(목표 너비 px, 목표 높이 px) → JPEG 바이트. 사진을 칸 크기에 맞춰 줄이는 함수</param>
    /// <returns>바꿨으면 true, 그 칸에 그림이 없으면 false</returns>
    public static bool ReplacePictureAt(WorksheetPart wsPart, string cellAddress,
                                        Func<int, int, byte[]> loadImage)
    {
        var anchor = FindPictureAt(wsPart, cellAddress);
        var blip   = anchor?.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
        if (anchor == null || blip == null || wsPart.DrawingsPart == null) return false;

        // 인쇄 품질을 위해 화면 크기(96 DPI)의 2배 해상도로 만든다
        var (cx, cy) = GetPictureSize(anchor);
        int w = Math.Max(1, (int)Math.Round(cx / EmuPerPixel * 2));
        int h = Math.Max(1, (int)Math.Round(cy / EmuPerPixel * 2));
        byte[] bytes = loadImage(w, h);

        var newPart = wsPart.DrawingsPart.AddImagePart(ImagePartType.Jpeg);
        using (var ms = new MemoryStream(bytes))
            newPart.FeedData(ms);

        string? oldId = blip.Embed?.Value;
        blip.Embed = wsPart.DrawingsPart.GetIdOfPart(newPart);

        // 앞 날짜 파일을 바탕으로 만들면 어제 사진이 파일 안에 남아 매일 커진다.
        // 다른 그림이 같이 쓰지 않는 옛 이미지면 지운다.
        if (!string.IsNullOrEmpty(oldId)
            && !wsPart.DrawingsPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
                     .Any(b => b.Embed?.Value == oldId))
        {
            try { wsPart.DrawingsPart.DeletePart(oldId); } catch { /* 못 지우면 그대로 둔다 */ }
        }
        return true;
    }

    // ── 그림·글상자 옮기기 ───────────────────────────────────────────

    /// <summary>
    /// 저장된 위치(preset)대로 그림·글상자들을 옮긴다. 이름이 시트에 없으면 건너뛴다.
    /// </summary>
    /// <returns>옮긴 개수</returns>
    public static int ApplyPreset(WorksheetPart wsPart, ShapePreset preset)
    {
        int moved = 0;
        var anchors = Anchors(wsPart).ToList();

        foreach (var p in preset.Shapes)
        {
            // 순서가 있으면 순서로 (시트마다 이름이 달라도 같은 그림), 없으면 이름으로
            var anchor = p.Index >= 0 && p.Index < anchors.Count ? anchors[p.Index] : FindByName(anchors, p.Name);
            if (anchor == null) continue;

            if (From(anchor) is { } f)
                SetMarker(f, p.FromCol, p.FromColOff, p.FromRow, p.FromRowOff);
            // oneCellAnchor 는 to 가 없고 크기(ext)로 정해지므로 from 만 바꾸면 된다
            if (anchor is Xdr.TwoCellAnchor && To(anchor) is { } t)
                SetMarker(t, p.ToCol, p.ToColOff, p.ToRow, p.ToRowOff);
            moved++;
        }
        return moved;
    }

    /// <summary>현재 파일에 있는 그림·글상자의 위치를 읽는다 (설정 확인·테스트용).</summary>
    public static ShapePlacement? ReadPlacement(WorksheetPart wsPart, string shapeName)
    {
        var a = FindByName(Anchors(wsPart).ToList(), shapeName);
        if (a == null || From(a) is not { } f) return null;
        var t = To(a);
        return new ShapePlacement
        {
            Name = shapeName,
            FromCol = Int(f.ColumnId), FromColOff = Long(f.ColumnOffset),
            FromRow = Int(f.RowId),    FromRowOff = Long(f.RowOffset),
            ToCol = Int(t?.ColumnId),  ToColOff = Long(t?.ColumnOffset),
            ToRow = Int(t?.RowId),     ToRowOff = Long(t?.RowOffset),
        };
    }

    private static void SetMarker(OpenXmlCompositeElement marker, int col, long colOff, int row, long rowOff)
    {
        var colId  = marker.GetFirstChild<Xdr.ColumnId>();
        var colOfs = marker.GetFirstChild<Xdr.ColumnOffset>();
        var rowId  = marker.GetFirstChild<Xdr.RowId>();
        var rowOfs = marker.GetFirstChild<Xdr.RowOffset>();
        if (colId  != null) colId.Text  = col.ToString();
        if (colOfs != null) colOfs.Text = colOff.ToString();
        if (rowId  != null) rowId.Text  = row.ToString();
        if (rowOfs != null) rowOfs.Text = rowOff.ToString();
    }
}
