namespace FormFiller.Core.Models;

/// <summary>
/// 블록 하나의 오늘 값 (DailyInput.Blocks 에 블록 이름으로 들어간다).
/// 블록 종류에 따라 쓰는 칸만 쓴다: value → Text, lines → Lines, counter → Prev/Today,
/// table → Rows, roster → People.
/// </summary>
public class BlockValue
{
    /// <summary>value: 입력한 글자</summary>
    public string Text { get; set; } = "";

    /// <summary>lines: 줄 목록</summary>
    public List<string> Lines { get; set; } = [];

    /// <summary>counter: 어제까지의 누계 (어제 파일에서 읽음, 화면에서 고칠 수 있음)</summary>
    public decimal Prev { get; set; }

    /// <summary>counter: 오늘 값</summary>
    public decimal Today { get; set; }

    /// <summary>table: 줄 목록</summary>
    public List<TableRowValue> Rows { get; set; } = [];

    /// <summary>roster: 오늘 나온 사람 (적힌 순서대로 표에 쓴다)</summary>
    public List<RosterEntry> People { get; set; } = [];
}

/// <summary>표 블록의 줄 하나</summary>
public class TableRowValue
{
    public string Name { get; set; } = "";
    public string Spec { get; set; } = "";
    /// <summary>전일 (어제까지 누계)</summary>
    public decimal Prev { get; set; }
    /// <summary>금일</summary>
    public decimal Today { get; set; }
    /// <summary>어제 파일의 금일 ([전일과 같게] 버튼용, 파일에는 쓰지 않음)</summary>
    public decimal LastToday { get; set; }
}

/// <summary>명단 블록의 사람 한 명</summary>
public class RosterEntry
{
    public string Job    { get; set; } = "";
    public string Person { get; set; } = "";
    public string Task   { get; set; } = "";
}
