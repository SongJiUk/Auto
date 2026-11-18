using System.IO.Compression;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FormFiller.Core.Config;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace FormFiller.Core.Excel;

/// <summary>
/// 저장 후 자동 검증 (CLAUDE.md "검증 규칙").
/// 만든 파일을 다시 열어, 바탕 파일(서식·앞 날짜 파일)과 비교해 "엑셀에서 깨져 보일 만한 것"을 찾는다.
///
/// [검사 항목]
/// 1. 파일이 열리는지 (zip·XML 이 망가지지 않았는지)
/// 2. 스키마 검사: 바탕 파일에 없던 오류가 새로 생겼는지 (실제 서식은 원래 오류가 조금 있을 수 있어 "늘었는지"만 본다)
/// 3. 시트마다 그림·도형 개수가 같은지, 그림 위치가 그대로인지 (그림 위치 항목으로 옮긴 그림은 빼고)
/// 4. 그림이 가리키는 이미지가 파일 안에 실제로 있는지 (없으면 엑셀에 빨간 X)
/// hwpx: zip 이 열리는지, mimetype 이 맨 앞·무압축인지, 본문 XML 이 올바른지
///
/// 문제를 찾아도 파일은 그대로 둔다. 화면이 주황색으로 알려 사람이 엑셀에서 열어 확인하게 한다.
/// </summary>
public static class OutputVerifier
{
    private const int MaxMessages = 5;

    public static List<string> Verify(string basePath, string outputPath, DocumentConfig doc)
    {
        try
        {
            return HwpxWriter.IsHwpx(outputPath) ? VerifyHwpx(outputPath) : VerifyXlsx(basePath, outputPath, doc);
        }
        catch (Exception ex)
        {
            // 검사기 자체 문제로 생성을 실패시키지는 않는다
            AppLog.Error($"검증 중 오류: {outputPath}", ex);
            return [$"자동 검증을 끝내지 못했습니다 ({ex.Message}) — 엑셀에서 열어 확인해 주세요"];
        }
    }

    // ── xlsx ─────────────────────────────────────────────────────────

    private static List<string> VerifyXlsx(string basePath, string outputPath, DocumentConfig doc)
    {
        var problems = new List<string>();

        SpreadsheetDocument output;
        try { output = SpreadsheetDocument.Open(outputPath, isEditable: false); }
        catch (Exception ex) { return [$"만든 파일이 열리지 않습니다: {ex.Message}"]; }

        using (output)
        using (var source = SpreadsheetDocument.Open(basePath, isEditable: false))
        {
            // 2. 스키마 검사: 새로 생긴 오류만
            var before = ErrorCounts(source);
            foreach (var (key, count) in ErrorCounts(output))
                if (count > before.GetValueOrDefault(key))
                    problems.Add($"파일 구조 오류가 생겼습니다 ({key.Part}): {key.Description}");

            // 3·4. 시트마다 그림
            var movable = MovableShapeIndexes(doc);
            foreach (string sheet in doc.TargetSheets())
            {
                var srcWs = CellHelper.FindWorksheetPart(source.WorkbookPart!, sheet);
                var outWs = CellHelper.FindWorksheetPart(output.WorkbookPart!, sheet);
                if (srcWs == null || outWs == null) continue;
                CompareDrawings(sheet, srcWs, outWs, movable, problems);
            }
        }

        return problems.Count <= MaxMessages ? problems
            : [.. problems.Take(MaxMessages), $"…그 밖에 {problems.Count - MaxMessages}건"];
    }

    private static Dictionary<(string Part, string Description), int> ErrorCounts(SpreadsheetDocument doc)
    {
        var counts = new Dictionary<(string, string), int>();
        foreach (var e in new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc))
        {
            var key = (e.Part?.Uri.ToString() ?? "", e.Description ?? "");
            counts[key] = counts.GetValueOrDefault(key) + 1;
        }
        return counts;
    }

    /// <summary>"그림 위치" 항목으로 일부러 옮기는 그림의 순서 (이 그림들은 위치가 바뀌어도 정상)</summary>
    private static HashSet<int> MovableShapeIndexes(DocumentConfig doc) =>
        doc.Fields.Where(f => f.Type == FieldTypes.ShapePosition)
           .SelectMany(f => (f.ShapeIndexes ?? []).Concat(f.Presets?.SelectMany(p => p.Shapes.Select(s => s.Index)) ?? []))
           .Where(i => i >= 0)
           .ToHashSet();

    private static List<OpenXmlCompositeElement> Anchors(WorksheetPart ws) =>
        ws.DrawingsPart?.WorksheetDrawing?.ChildElements.OfType<OpenXmlCompositeElement>()
          .Where(e => e is Xdr.TwoCellAnchor or Xdr.OneCellAnchor or Xdr.AbsoluteAnchor).ToList()
        ?? [];

    private static string Position(OpenXmlCompositeElement a)
    {
        static string M(OpenXmlCompositeElement? m) => m == null ? "-" :
            $"{m.GetFirstChild<Xdr.ColumnId>()?.Text},{m.GetFirstChild<Xdr.ColumnOffset>()?.Text}," +
            $"{m.GetFirstChild<Xdr.RowId>()?.Text},{m.GetFirstChild<Xdr.RowOffset>()?.Text}";
        return M(a.GetFirstChild<Xdr.FromMarker>()) + "|" + M(a.GetFirstChild<Xdr.ToMarker>());
    }

    private static void CompareDrawings(string sheet, WorksheetPart src, WorksheetPart output,
                                        HashSet<int> movable, List<string> problems)
    {
        var a = Anchors(src);
        var b = Anchors(output);
        if (a.Count != b.Count)
        {
            problems.Add($"[{sheet}] 그림·도형 개수가 달라졌습니다 ({a.Count}개 → {b.Count}개)");
            return;   // 개수가 다르면 순서 비교는 의미 없다
        }

        int srcPics = a.Count(x => x.Descendants<Xdr.Picture>().Any());
        int outPics = b.Count(x => x.Descendants<Xdr.Picture>().Any());
        if (srcPics != outPics)
            problems.Add($"[{sheet}] 그림 개수가 달라졌습니다 ({srcPics}개 → {outPics}개)");

        for (int i = 0; i < a.Count; i++)
            if (!movable.Contains(i) && Position(a[i]) != Position(b[i]))
                problems.Add($"[{sheet}] '{ExcelDrawingHelper.ShapeName(b[i]) ?? $"{i + 1}번째 그림"}' 위치가 바뀌었습니다");

        // 그림이 가리키는 이미지가 파일 안에 있는지
        if (output.DrawingsPart is { } dp)
            foreach (var blip in dp.WorksheetDrawing.Descendants<A.Blip>())
            {
                string? id = blip.Embed?.Value;
                if (string.IsNullOrEmpty(id)) continue;
                try { _ = dp.GetPartById(id); }
                catch (ArgumentOutOfRangeException) { problems.Add($"[{sheet}] 그림의 이미지가 파일에 없습니다 (엑셀에서 빨간 X)"); }
            }
    }

    // ── hwpx ─────────────────────────────────────────────────────────

    private static List<string> VerifyHwpx(string outputPath)
    {
        var problems = new List<string>();
        ZipArchive zip;
        try { zip = ZipFile.OpenRead(outputPath); }
        catch (Exception ex) { return [$"만든 파일이 열리지 않습니다: {ex.Message}"]; }

        using (zip)
        {
            var first = zip.Entries.FirstOrDefault();
            if (first?.FullName != "mimetype")
                problems.Add("mimetype 이 맨 앞에 없습니다 (한글에서 안 열릴 수 있음)");
            else if (first.CompressedLength != first.Length)
                problems.Add("mimetype 이 압축되어 있습니다 (한글에서 안 열릴 수 있음)");

            foreach (var e in zip.Entries.Where(e => e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
            {
                try
                {
                    using var s = e.Open();
                    new XmlDocument().Load(s);
                }
                catch (XmlException ex) { problems.Add($"{e.FullName} 내용이 망가졌습니다: {ex.Message}"); }
            }
        }
        return problems;
    }
}
