using FormFiller.Core.Config;

namespace FormFiller.Core.Excel;

/// <summary>
/// 표·명단의 머리글(예: "장비명 · 규격 · 전일(대) · 금일(대) · 누계(대)")을 읽어 열 역할을 자동으로 정한다.
/// 사용자는 행만 드래그하면 되고, 찾은 열이 다를 때만 다시 클릭한다.
///
/// [찾는 방법]
/// 데이터 첫 행 바로 위 1~3행 중, 역할이 2개 이상 맞는 가장 가까운 행을 머리글로 본다.
/// 같은 역할이 여러 열에 있으면 왼쪽 것을 쓴다 (선택한 범위 안에서만 찾는다).
/// 칸 글자는 함수로 받는다 → 화면은 Excel 에게 묻고, 테스트는 파일에서 읽는다.
/// </summary>
public static class HeaderDetector
{
    /// <param name="kind">BlockTypes.Table 또는 BlockTypes.Roster</param>
    /// <param name="startRow">데이터 첫 행 (1부터)</param>
    /// <param name="firstCol">선택 범위 첫 열 (1부터)</param>
    /// <param name="lastCol">선택 범위 끝 열. 한 열만 드래그했으면 오른쪽으로 조금 더 본다.</param>
    /// <param name="cellText">(행, 열) → 칸에 보이는 글자</param>
    /// <returns>역할 → 열 문자 (예: "today" → "H")</returns>
    public static Dictionary<string, string> Detect(string kind, int startRow, int firstCol, int lastCol,
                                                    Func<int, int, string?> cellText)
    {
        if (lastCol <= firstCol) lastCol = firstCol + 12;   // 한 열만 골랐으면 옆 열들도 본다

        Dictionary<string, string>? best = null;
        for (int row = startRow - 1; row >= Math.Max(1, startRow - 3); row--)
        {
            var found = new Dictionary<string, string>();
            for (int col = firstCol; col <= lastCol; col++)
            {
                string? role = RoleFor(kind, cellText(row, col));
                if (role != null) found.TryAdd(role, ColLetter(col));   // 왼쪽 것 먼저
            }
            if (found.Count >= 2) { best = found; break; }
        }
        if (best != null && kind == BlockTypes.Table) TreatPrevAsTodayWhenNoToday(best);
        return best ?? [];
    }

    /// <summary>
    /// 금일 열 없이 전일·누계만 있는 표 (폐기물 "전일반출 · 누계", 철거 "전일철거동 · 누계동"):
    /// 전일 칸이 매일 적는 칸이고 누계 = 어제 누계 + 그 값 → 전일 열을 금일 역할로 바꾼다 (2026-11-22 사용자 규칙).
    /// </summary>
    private static void TreatPrevAsTodayWhenNoToday(Dictionary<string, string> found)
    {
        if (found.ContainsKey(BlockTypes.RoleToday) || !found.ContainsKey(BlockTypes.RoleTotal)) return;
        if (!found.Remove(BlockTypes.RolePrev, out string? col)) return;
        found[BlockTypes.RoleToday] = col;
    }

    /// <summary>머리글 글자 하나의 역할. 모르면 null.</summary>
    public static string? RoleFor(string kind, string? header)
    {
        if (string.IsNullOrWhiteSpace(header)) return null;
        string h = new(header.Where(c => !char.IsWhiteSpace(c)).ToArray());   // "직 종" → "직종", "총 누계(명)" → "총누계(명)"

        if (kind == BlockTypes.Roster)
        {
            if (Has(h, "작업내용", "작업", "업무", "담당")) return BlockTypes.RoleTask;
            if (Has(h, "성명", "이름", "근로자", "작업자")) return BlockTypes.RolePerson;
            if (Has(h, "직종", "직위", "직책", "직급", "공종")) return BlockTypes.RoleJob;
            return null;
        }

        // 표: "전일누계" 는 전일, "총누계" 는 누계 → 전일을 먼저 본다
        if (Has(h, "전일", "어제", "이전")) return BlockTypes.RolePrev;
        if (Has(h, "금일", "당일", "오늘")) return BlockTypes.RoleToday;
        if (Has(h, "누계", "합계", "총계")) return BlockTypes.RoleTotal;
        if (Has(h, "규격", "사양", "용량", "톤수")) return BlockTypes.RoleSpec;
        if (Has(h, "장비명", "장비", "직종", "공종", "품명", "명칭", "이름", "구분", "자재", "항목")) return BlockTypes.RoleName;
        return null;
    }

    private static bool Has(string h, params string[] words) => words.Any(w => h.Contains(w, StringComparison.OrdinalIgnoreCase));

    /// <summary>1 → "A", 27 → "AA"</summary>
    public static string ColLetter(int col)
    {
        string r = "";
        while (col > 0) { col--; r = (char)('A' + col % 26) + r; col /= 26; }
        return r;
    }

    /// <summary>파일의 한 시트에서 칸 글자를 읽는 함수 (테스트·화면에서 Excel 없이 쓸 때)</summary>
    public static Func<int, int, string?> FromFile(DocumentFormat.OpenXml.Packaging.WorksheetPart ws) =>
        (row, col) => CellHelper.GetCellValue(ws, ColLetter(col) + row);
}
