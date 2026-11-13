using System.IO.Compression;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using FormFiller.Core.Config;
using FormFiller.Core.Models;

namespace FormFiller.Core.Excel;

/// <summary>
/// 한글(hwpx) 서식 채우기.
///
/// [원리]
/// hwpx 는 xlsx 처럼 zip 파일이다. 본문 글자는 Contents/section0.xml 등에 들어 있다.
/// 사용자가 한글에서 서식의 채울 자리에 {{날짜}} 같은 "자리 표시"를 적어 두면,
/// 생성할 때 그 글자만 값으로 바꾼다. 표·글꼴·그림 등 나머지는 건드리지 않는다.
///
/// [주의]
/// - 예전 .hwp(바이너리) 파일은 지원하지 않는다. 한글에서 "다른 이름으로 저장 → hwpx" 로 바꿔 쓴다.
/// - 자리 표시 글자 중간에 글꼴·색을 바꾸면 내부에서 글자가 쪼개져 찾지 못한다. 한 번에 입력해 둔다.
/// - zip 안의 "mimetype" 항목은 반드시 맨 앞에, 압축하지 않은 상태로 있어야 한글이 연다.
/// </summary>
public static class HwpxWriter
{
    private static readonly Regex PlaceholderRegex = new(@"\{\{[^{}<>]{1,40}\}\}", RegexOptions.Compiled);

    public static bool IsHwpx(string path) =>
        path.EndsWith(".hwpx", StringComparison.OrdinalIgnoreCase);

    /// <summary>필드의 자리 표시 글자. 따로 정하지 않았으면 {{필드 이름}}.</summary>
    public static string PlaceholderOf(FieldConfig f) =>
        string.IsNullOrWhiteSpace(f.Placeholder) ? "{{" + f.Label + "}}" : f.Placeholder!;

    /// <summary>서식 파일 본문에 적힌 자리 표시 목록 (예: ["{{날짜}}", "{{날씨}}"])</summary>
    public static List<string> FindPlaceholders(string hwpxPath)
    {
        var found = new List<string>();
        using var zip = ZipFile.OpenRead(hwpxPath);
        foreach (var e in zip.Entries.Where(IsContentXml))
        {
            using var r = new StreamReader(e.Open(), Encoding.UTF8);
            foreach (Match m in PlaceholderRegex.Matches(r.ReadToEnd()))
                if (!found.Contains(m.Value)) found.Add(m.Value);
        }
        return found;
    }

    /// <summary>
    /// 서식을 복사해 자리 표시를 값으로 바꾼 hwpx 를 만든다. 원본은 읽기만 한다.
    /// </summary>
    public static void Generate(string templatePath, string outputPath, DocumentConfig docConfig,
                                DailyInput input, List<string>? notes = null)
    {
        // 자리 표시 → 바꿀 값
        string sheet = docConfig.TargetSheets().FirstOrDefault() ?? "";
        var values = new Dictionary<string, string>();
        foreach (var f in docConfig.Fields)
        {
            string? v = GenericExcelWriter.GetSimpleValue(f, input, docConfig.Id, sheet);
            if (v != null) values[PlaceholderOf(f)] = v;
        }

        var replaced = new HashSet<string>();
        string tmp = outputPath + ".tmp";

        using (var src = ZipFile.OpenRead(templatePath))
        using (var dst = ZipFile.Open(tmp, ZipArchiveMode.Create))
        {
            // mimetype 을 맨 앞에, 압축 없이
            var mime = src.GetEntry("mimetype");
            if (mime != null) CopyEntry(mime, dst, CompressionLevel.NoCompression);

            foreach (var e in src.Entries.Where(e => e.FullName != "mimetype"))
            {
                if (!IsContentXml(e)) { CopyEntry(e, dst, CompressionLevel.Optimal); continue; }

                string xml;
                using (var r = new StreamReader(e.Open(), Encoding.UTF8)) xml = r.ReadToEnd();

                foreach (var (ph, v) in values)
                {
                    if (!xml.Contains(ph)) continue;
                    xml = xml.Replace(ph, ToHwpxText(v));
                    replaced.Add(ph);
                }

                var ne = dst.CreateEntry(e.FullName, CompressionLevel.Optimal);
                ne.LastWriteTime = e.LastWriteTime;
                using var w = new StreamWriter(ne.Open(), new UTF8Encoding(false));
                w.Write(xml);
            }
        }

        File.Move(tmp, outputPath, overwrite: true);

        foreach (var ph in values.Keys.Where(k => !replaced.Contains(k)))
            notes?.Add($"{docConfig.Name}: 서식에서 {ph} 자리를 찾지 못함");
    }

    // ── 내부 ─────────────────────────────────────────────────────────

    /// <summary>본문·머리말 등 글자가 들어 있는 XML (Contents 폴더)</summary>
    private static bool IsContentXml(ZipArchiveEntry e) =>
        e.FullName.StartsWith("Contents/", StringComparison.OrdinalIgnoreCase)
        && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// 값을 hwpx 본문에 넣을 수 있게 바꾼다: XML 특수문자(&amp;, &lt; 등) 처리,
    /// 줄바꿈은 한글의 줄 나눔 표시(&lt;hp:lineBreak/&gt;)로.
    /// </summary>
    private static string ToHwpxText(string value)
    {
        var lines = value.Replace("\r\n", "\n").Split('\n').Select(l => SecurityElement.Escape(l) ?? "");
        return string.Join("<hp:lineBreak/>", lines);
    }

    private static void CopyEntry(ZipArchiveEntry e, ZipArchive dst, CompressionLevel level)
    {
        var ne = dst.CreateEntry(e.FullName, level);
        ne.LastWriteTime = e.LastWriteTime;
        using var s = e.Open();
        using var d = ne.Open();
        s.CopyTo(d);
    }
}
