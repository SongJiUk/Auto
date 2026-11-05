using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace FormFiller.Core.Excel;

/// <summary>
/// xlsx 파일의 특정 시트에 있는 이미지 슬롯을 교체한다.
///
/// [원리]
/// xlsx는 zip 파일이다.
///   xl/media/imageN.xxx  — 실제 이미지 바이너리
///   xl/drawings/drawingN.xml — 시트 위 이미지 위치(앵커) 정보
///
/// 이미 위치가 잡혀 있는 서식 파일에서 이미지 바이너리만 교체하면
/// 셀 크기·앵커·레이아웃은 그대로 유지된다.
/// 서식 파일의 drawing XML에 이미 &lt;a:stretch&gt;&lt;a:fillRect/&gt;가 있으므로
/// 교체된 이미지는 슬롯 크기에 맞게 자동으로 채워진다(가득채우기).
/// </summary>
public static class ExcelPhotoWriter
{
    // ════════════════════════════════════════════════════════════════
    // ── 슬롯 정보 ────────────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    /// <summary>이미지 슬롯 하나의 크기 정보</summary>
    public record PhotoSlotInfo(int SlotIndex, long WidthEmu, long HeightEmu)
    {
        private const double EmuPerPx = 914400.0 / 96.0; // 96 DPI 기준

        /// <summary>96 DPI 기준 픽셀 너비</summary>
        public int WidthPx  => (int)Math.Round(WidthEmu  / EmuPerPx);

        /// <summary>96 DPI 기준 픽셀 높이</summary>
        public int HeightPx => (int)Math.Round(HeightEmu / EmuPerPx);
    }

    /// <summary>
    /// 지정 시트에 있는 모든 이미지 슬롯의 크기 정보를 반환한다.
    /// 빈 목록이면 해당 시트에 이미지가 없다.
    /// </summary>
    public static List<PhotoSlotInfo> GetPhotoSlotInfos(string xlsxPath, string sheetName)
    {
        var result = new List<PhotoSlotInfo>();
        if (!File.Exists(xlsxPath)) return result;

        try
        {
            using var doc = SpreadsheetDocument.Open(xlsxPath, isEditable: false);
            var wsPart = GetWorksheetPart(doc, sheetName);
            if (wsPart?.DrawingsPart == null) return result;

            int idx = 0;
            foreach (var anchor in wsPart.DrawingsPart.WorksheetDrawing
                                         .Elements<Xdr.TwoCellAnchor>())
            {
                // 그림(Picture) 요소가 있는 앵커만 슬롯으로 처리한다
                var pic = anchor.Descendants<Xdr.Picture>().FirstOrDefault();
                if (pic == null) { idx++; continue; }

                // spPr > xfrm > ext 에서 EMU 단위 크기 읽기
                var ext = pic.ShapeProperties
                             ?.Transform2D
                             ?.Extents;

                long cx = ext?.Cx?.Value ?? 0;
                long cy = ext?.Cy?.Value ?? 0;
                result.Add(new PhotoSlotInfo(idx, cx, cy));
                idx++;
            }
        }
        catch { /* 파일 열기 실패 시 빈 목록 반환 */ }

        return result;
    }

    // ════════════════════════════════════════════════════════════════
    // ── 슬롯 이미지 읽기 ─────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    /// <summary>
    /// 지정 시트의 슬롯 이미지 바이너리를 반환한다.
    /// 서식 파일에 들어 있는 플레이스홀더 이미지를 미리보기용으로 읽을 때 사용한다.
    /// 슬롯이 없거나 오류 시 null 반환.
    /// </summary>
    public static byte[]? GetSlotImageBytes(string xlsxPath, string sheetName, int slotIndex)
    {
        if (!File.Exists(xlsxPath)) return null;

        try
        {
            using var doc = SpreadsheetDocument.Open(xlsxPath, isEditable: false);
            var wsPart = GetWorksheetPart(doc, sheetName);
            if (wsPart?.DrawingsPart == null) return null;

            var pictureAnchors = wsPart.DrawingsPart.WorksheetDrawing
                .Elements<Xdr.TwoCellAnchor>()
                .Where(a => a.Descendants<Xdr.Picture>().Any())
                .ToList();

            if (slotIndex >= pictureAnchors.Count) return null;

            var pic = pictureAnchors[slotIndex].Descendants<Xdr.Picture>().First();
            var rId = pic.BlipFill?.Blip?.Embed?.Value;
            if (rId == null) return null;

            var imagePart = (ImagePart)wsPart.DrawingsPart.GetPartById(rId);
            using var ms = new MemoryStream();
            imagePart.GetStream().CopyTo(ms);
            return ms.ToArray();
        }
        catch { return null; }
    }

    // ════════════════════════════════════════════════════════════════
    // ── 이미지 교체 ──────────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    /// <summary>
    /// xlsx 파일의 지정 시트에서 슬롯 번호(0-based) 위치의 이미지를
    /// imageBytes(JPEG 또는 PNG 바이너리)로 교체한다.
    ///
    /// "가득채우기" 적용 방법:
    ///   호출 전에 PhotoService.ResizeToFill()로 이미지를 슬롯 크기에 맞게
    ///   리사이즈한 후 전달하면 슬롯을 정확히 채운다.
    ///   drawing XML에 이미 &lt;a:stretch&gt;가 있으므로 Excel에서도 꽉 찬다.
    /// </summary>
    /// <param name="xlsxPath">교체할 xlsx (복사본에 대해 호출)</param>
    /// <param name="sheetName">시트 이름</param>
    /// <param name="slotIndex">교체할 이미지 순서 (0-based)</param>
    /// <param name="imageBytes">새 이미지 바이너리 (JPEG 권장)</param>
    public static void ReplacePhoto(
        string xlsxPath, string sheetName, int slotIndex, byte[] imageBytes)
    {
        if (!File.Exists(xlsxPath))
            throw new FileNotFoundException("xlsx 파일을 찾을 수 없습니다.", xlsxPath);

        using var doc = SpreadsheetDocument.Open(xlsxPath, isEditable: true);
        var wsPart = GetWorksheetPart(doc, sheetName)
            ?? throw new ArgumentException($"시트를 찾을 수 없습니다: {sheetName}");

        var drawingPart = wsPart.DrawingsPart
            ?? throw new InvalidOperationException($"시트 '{sheetName}'에 그림이 없습니다.");

        // 그림(Picture) 앵커만 순서대로 나열
        var pictureAnchors = drawingPart.WorksheetDrawing
            .Elements<Xdr.TwoCellAnchor>()
            .Where(a => a.Descendants<Xdr.Picture>().Any())
            .ToList();

        if (slotIndex >= pictureAnchors.Count)
            throw new ArgumentOutOfRangeException(nameof(slotIndex),
                $"슬롯 {slotIndex}이 없습니다. 이 시트의 이미지 수: {pictureAnchors.Count}");

        // 해당 슬롯의 rId 추출
        var pic  = pictureAnchors[slotIndex].Descendants<Xdr.Picture>().First();
        var rId  = pic.BlipFill?.Blip?.Embed?.Value
            ?? throw new InvalidOperationException("이미지 관계 ID를 찾을 수 없습니다.");

        // ImagePart 데이터 교체
        var imagePart = (ImagePart)drawingPart.GetPartById(rId);
        using var ms  = new MemoryStream(imageBytes);
        imagePart.FeedData(ms);
    }

    // ════════════════════════════════════════════════════════════════
    // ── 헬퍼 ─────────────────────────────────────────────────────────
    // ════════════════════════════════════════════════════════════════

    /// <summary>워크북에서 시트 이름으로 WorksheetPart를 찾는다.</summary>
    private static WorksheetPart? GetWorksheetPart(SpreadsheetDocument doc, string sheetName)
    {
        var sheet = doc.WorkbookPart?.Workbook.Descendants<Sheet>()
            .FirstOrDefault(s => s.Name?.Value == sheetName);
        if (sheet == null) return null;

        return doc.WorkbookPart!.GetPartById(sheet.Id!.Value!) as WorksheetPart;
    }
}
