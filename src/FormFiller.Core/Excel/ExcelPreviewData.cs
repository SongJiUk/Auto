namespace FormFiller.Core.Excel;

/// <summary>엑셀 셀 가로 정렬 (HorizontalAlignmentValues 를 단순화한 열거형)</summary>
public enum ExcelHAlign
{
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
}

/// <summary>엑셀 셀 테두리 (상·하·좌·우 존재 여부)</summary>
public readonly struct ExcelBorderInfo
{
    public bool Left   { get; init; }
    public bool Top    { get; init; }
    public bool Right  { get; init; }
    public bool Bottom { get; init; }
}

/// <summary>
/// 병합 셀 범위 (1-based 행·열).
/// TopRow/LeftCol 이 병합의 시작(대표 셀), BottomRow/RightCol 이 끝이다.
/// </summary>
public sealed class ExcelMergeRange
{
    public int TopRow    { get; init; }  // 1-based
    public int LeftCol   { get; init; }  // 1-based
    public int BottomRow { get; init; }  // 1-based
    public int RightCol  { get; init; }  // 1-based

    public int RowSpan => BottomRow - TopRow + 1;
    public int ColSpan => RightCol  - LeftCol  + 1;
}

/// <summary>
/// ExcelAnalyzer.GetPreviewData() 가 반환하는 미리보기 데이터 묶음.
/// View(code-behind) 에서 이 객체 하나만 받아 Grid 를 완성한다.
///
/// 모든 배열은 0-based: [row, col] 에서 row=0 이 1행, col=0 이 A열.
/// ColWidthsPx / RowHeightsPx 도 0-based: index=0 이 첫 열/행.
/// </summary>
public sealed class ExcelPreviewData
{
    public int Rows { get; init; }
    public int Cols { get; init; }

    public string[,]          Values       { get; init; } = new string[0, 0];
    public uint?[,]           BgColors     { get; init; } = new uint?[0, 0];
    public ExcelHAlign[,]     HAlignments  { get; init; } = new ExcelHAlign[0, 0];
    public bool[,]            IsBold       { get; init; } = new bool[0, 0];
    public ExcelBorderInfo[,] Borders      { get; init; } = new ExcelBorderInfo[0, 0];

    /// <summary>각 열의 픽셀 너비 (0-based, length == Cols)</summary>
    public double[] ColWidthsPx  { get; init; } = [];

    /// <summary>각 행의 픽셀 높이 (0-based, length == Rows)</summary>
    public double[] RowHeightsPx { get; init; } = [];

    /// <summary>병합 셀 목록</summary>
    public List<ExcelMergeRange> Merges { get; init; } = [];
}
