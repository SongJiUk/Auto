using FormFiller.Core.Config;
using FormFiller.Core.Models;

namespace FormFiller.Tests.Helpers;

/// <summary>
/// 공사일보 모양을 흉내 낸 가짜 서식 + 설정.
/// 칸 배치는 docs/샘플_셀지도.md 의 공사일보와 같은 방식(장비 A/D/F/H/J, 인원 L/N/P/R)이다.
///
///   장비(20~24행): A 장비명, D 규격, F 전일, H 금일, J 누계(=F+H)
///   인원(20~24행): L 직종,   N 전일누계, P 금일, R 총누계(=N+P)
///   25행: 합계 수식
/// </summary>
public static class DailyReportSample
{
    public const string Sheet = "공사일보";
    public const string TemplateName = "서식_공사일보.xlsx";
    public const string Pattern = "{yy}.{MM}.{dd} 공사일보.xlsx";

    public static string CreateTemplate(string path)
    {
        var s = new SheetSpec(Sheet)
        {
            ["C6"] = "2026년 11월 04일 수요일",
            ["L6"] = "맑음",
            ["A9"] = "어제 작업1", ["A10"] = "어제 작업2", ["A11"] = "어제 작업3",
            ["L9"] = "내일 작업1",

            // 장비
            ["A20"] = "B/H 380", ["D20"] = "39.6t", ["F20"] = 10, ["H20"] = 1, ["J20"] = new F("F20+H20", "11"),
            ["A21"] = "덤프",    ["D21"] = "15m3",  ["F21"] = 5,  ["H21"] = 2, ["J21"] = new F("F21+H21", "7"),
            ["A22"] = "세륜기",                     ["F22"] = 3,  ["H22"] = 1, ["J22"] = 4,   // 누계가 수식이 아니라 숫자
            ["J23"] = new F("F23+H23", "0"),
            ["J24"] = new F("F24+H24", "0"),
            ["F25"] = new F("SUM(F20:F24)", "18"), ["H25"] = new F("SUM(H20:H24)", "4"),

            // 인원
            ["L20"] = "관리자", ["N20"] = 30, ["P20"] = 3, ["R20"] = new F("N20+P20", "33"),
            ["L21"] = "비계",   ["N21"] = 50, ["P21"] = 5, ["R21"] = new F("N21+P21", "55"),
            ["L22"] = "수장재", ["N22"] = 20, ["P22"] = 3, ["R22"] = new F("N22+P22", "23"),
            ["R23"] = new F("N23+P23", "0"),
            ["R24"] = new F("N24+P24", "0"),
            ["P25"] = new F("SUM(P20:P24)", "11"),
        };
        return XlsxBuilder.Create(path, s);
    }

    public static DocumentConfig Config(bool includesSujangjae = false, bool usePrevious = true) => new()
    {
        Id = "daily_report",
        Name = "공사일보",
        TemplateFile = TemplateName,
        OutputPattern = Pattern,
        SheetName = Sheet,
        IncludesSujangjae = includesSujangjae,
        UsePreviousAsBase = usePrevious,
        Fields =
        [
            new() { Id = "date",          Label = "날짜", Type = FieldTypes.DateKorean, Cell = "C6" },
            new() { Id = "weather",       Label = "날씨", Type = FieldTypes.Text,       Cell = "L6" },
            new() { Id = "today_work",    Label = "금일", Type = FieldTypes.Multiline,  Cells = ["A9", "A10", "A11"] },
            new() { Id = "tomorrow_work", Label = "명일", Type = FieldTypes.Multiline,  Cells = ["L9", "L10", "L11"] },
            new()
            {
                Id = "equipment_table", Label = "장비", Type = FieldTypes.EquipmentTable,
                StartRow = 20, EndRow = 24, NameCol = "A", SpecCol = "D", PrevDayCol = "F", TodayCol = "H",
            },
            new()
            {
                Id = "personnel_table", Label = "인원", Type = FieldTypes.PersonnelTable,
                StartRow = 20, EndRow = 24, RoleCol = "L", PrevCumulativeCol = "N", TodayCol = "P",
            },
        ],
    };

    /// <summary>화면에서 입력한 것과 같은 하루치 값. 전일 값은 비워 두고 LoadPreviousDay 로 채운다.</summary>
    public static DailyInput Input(DateTime date, int bh = 1, int dump = 2, int wash = 1,
                                   int manager = 3, int scaffold = 5, int sujang = 3) => new()
    {
        Date = date,
        Weather = "흐림",
        TodayMorningWork = "철거 작업",
        TodayAfternoonWork = "자재 반출",
        TomorrowWork = "현장 정리",
        Equipment =
        [
            new() { Name = "B/H 380", Spec = "39.6t", TodayCount = bh },
            new() { Name = "덤프",    Spec = "15m3",  TodayCount = dump },
            new() { Name = "세륜기",                  TodayCount = wash },
        ],
        Personnel =
        [
            new() { Role = "관리자", TodayCount = manager },
            new() { Role = "비계",   TodayCount = scaffold },
        ],
        SujangjaePersonnel = [new() { Role = "수장재", TodayCount = sujang }],
    };
}
