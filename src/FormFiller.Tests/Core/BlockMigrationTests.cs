using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;
using FormFiller.Core.Models;
using FormFiller.Tests.Helpers;

namespace FormFiller.Tests.Core;

/// <summary>
/// 예전 방식 → 블록 변환 테스트. 같은 입력이면 두 방식으로 만든 파일의 모든 칸 값이 같아야 한다.
/// </summary>
public class BlockMigrationTests : IDisposable
{
    private static readonly DateTime Thu = new(2026, 11, 5), Fri = new(2026, 11, 6);
    private readonly TempDir _legacyDir = new(), _blockDir = new();

    public void Dispose() { _legacyDir.Dispose(); _blockDir.Dispose(); }

    [Fact]
    public void 변환하면_예전_종류가_블록으로()
    {
        var blocks = BlockMigration.Convert(DailyReportSample.Config());
        blocks.Select(f => (f.Type, BlockTypes.KeyOf(f))).Should().Equal(
            (BlockTypes.Date, "날짜"), (BlockTypes.Value, "날씨"), (BlockTypes.Lines, "금일 작업"),
            (BlockTypes.Lines, "명일 작업"), (BlockTypes.Table, "장비"), (BlockTypes.Table, "인원"));
        var eq = blocks[4];
        BlockTypes.ColOf(eq, BlockTypes.RoleTotal).Should().Be("J", "예전 규칙: 금일 H + 2");
        BlockTypes.ColOf(eq, BlockTypes.RoleSpec).Should().Be("D");
        blocks[2].StartFrom.Should().Be("명일 작업");
        BlockMigration.HasLegacy(new DocumentConfig { Fields = blocks }).Should().BeFalse();
    }

    [Fact]
    public void TBM_인원은_인원_더하기_수장재_계산으로()
    {
        var tbm = new DocumentConfig
        {
            Fields =
            [
                new() { Id = "personnel", Label = "인원", Type = FieldTypes.TbmPersonnel, Cell = "AD6" },
                new() { Id = "rate", Label = "참석율", Type = FieldTypes.AttendanceRate, Cell = "U16" },
                new() { Id = "eq", Label = "장비", Type = FieldTypes.ActiveEquipmentNames, Cell = "F7" },
            ],
        };
        var blocks = BlockMigration.Convert(tbm);
        blocks[0].Sources.Should().Equal("인원", "수장재 인원");
        blocks[1].FixedValue.Should().Be("100");
        blocks[2].CalcKind.Should().Be(BlockTypes.CalcNames);
        blocks.Should().ContainSingle(f => f.Label == "수장재 인원" && f.Cell == null, "칸 없이 계산에만 쓰는 입력");
    }

    /// <summary>예전 방식 입력을 블록 입력으로 옮긴다 (화면이 할 일과 같다)</summary>
    private static void FillBlocks(DailyInput legacy, DailyInput target)
    {
        target.Blocks["날씨"].Text = legacy.Weather;
        target.Blocks["금일 작업"].Lines = [.. legacy.TodayWork.Split('\n')];
        target.Blocks["명일 작업"].Lines = [.. legacy.TomorrowWork.Split('\n')];
        foreach (var row in target.Blocks["장비"].Rows)
            row.Today = legacy.Equipment.FirstOrDefault(e => e.Name == row.Name)?.TodayCount ?? 0;
        foreach (var row in target.Blocks["인원"].Rows)
            row.Today = legacy.Personnel.FirstOrDefault(p => p.Role == row.Name)?.TodayCount ?? 0;
    }

    private static Dictionary<string, string?> AllCells(string path, string sheet)
    {
        using var r = new XlsxReader(path);
        return r.Sheet(sheet).Worksheet.Descendants<Cell>()
                .ToDictionary(c => c.CellReference!.Value!, c => r.Value(sheet, c.CellReference!.Value!));
    }

    [Fact]
    public void 두_방식으로_이틀_만든_공사일보의_모든_칸이_같다()
    {
        DailyReportSample.CreateTemplate(_legacyDir.File(DailyReportSample.TemplateName));
        DailyReportSample.CreateTemplate(_blockDir.File(DailyReportSample.TemplateName));
        var legacyCfg = DailyReportSample.Config();
        var blockCfg  = DailyReportSample.Config();
        blockCfg.Fields = BlockMigration.Convert(legacyCfg);

        foreach (var (date, bh, scaffold) in new[] { (Thu, 2, 6), (Fri, 4, 3) })
        {
            // 예전 방식
            var legacy = DailyReportSample.Input(date, bh: bh, scaffold: scaffold);
            var prevL = GenericExcelWriter.FindPreviousFile(date, legacyCfg.OutputPattern, _legacyDir.Path, "output");
            GenericExcelWriter.LoadPreviousDay(prevL?.Path ?? _legacyDir.File(legacyCfg.TemplateFile), legacy, legacyCfg);
            string lp = GenericExcelWriter.Generate(legacy, legacyCfg, _legacyDir.Path, "output");

            // 블록 방식 (같은 오늘 값)
            var prevB = GenericExcelWriter.FindPreviousFile(date, blockCfg.OutputPattern, _blockDir.Path, "output");
            var read  = BlockEngine.ReadValues(prevB?.Path ?? _blockDir.File(blockCfg.TemplateFile), blockCfg);
            var block = new DailyInput { Date = date, Blocks = BlockEngine.InitialValues(blockCfg, read) };
            FillBlocks(DailyReportSample.Input(date, bh: bh, scaffold: scaffold), block);
            string bp = GenericExcelWriter.Generate(block, blockCfg, _blockDir.Path, "output");

            AllCells(bp, DailyReportSample.Sheet).Should().BeEquivalentTo(AllCells(lp, DailyReportSample.Sheet),
                $"{date:MM/dd} 모든 칸 값이 같아야 한다");
        }
    }
}
