using System.Text.Json;
using FormFiller.Core.Config;
using FormFiller.Core.Models;

namespace FormFiller.Core.Excel;

/// <summary>
/// [🧪 시험 생성]: 설정을 저장하지 않고, 서식에서 새로 오늘 파일을 임시 폴더에 만들어 본다.
/// 입력값은 서식에서 읽은 처음 값(매일 입력 칸은 서식 글자, 누계는 0 더하기)이라 칸 위치·형식 확인용.
/// 실제 출력 폴더·원본 서식·넘겨받은 설정은 건드리지 않는다.
/// </summary>
public static class TrialGenerator
{
    private const string DefaultPattern = "{yy}.{MM}.{dd} 시험.xlsx";

    /// <param name="templatePath">서식 파일의 실제(절대) 경로</param>
    /// <param name="folder">시험 파일을 둘 폴더 (절대 경로)</param>
    /// <returns>만든 파일 경로</returns>
    public static string Generate(DocumentConfig doc, string templatePath, string folder, DateTime date,
                                  List<string>? notes = null)
    {
        var trial = JsonSerializer.Deserialize<DocumentConfig>(JsonSerializer.Serialize(doc))!;
        trial.TemplateFile      = templatePath;
        trial.UsePreviousAsBase = false;   // 늘 서식에서 새로
        if (string.IsNullOrWhiteSpace(trial.OutputPattern)) trial.OutputPattern = DefaultPattern;
        if (!trial.OutputPattern.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            trial.OutputPattern = Path.ChangeExtension(trial.OutputPattern, ".xlsx");

        var input = new DailyInput
        {
            Date   = date.Date,
            Blocks = BlockEngine.InitialValues(trial, BlockEngine.ReadValues(templatePath, trial)),
        };
        return GenericExcelWriter.Generate(input, trial, "", folder, notes: notes);
    }

    /// <summary>시험 파일용 새 임시 폴더 경로 (만들 때마다 시각별로 따로)</summary>
    public static string NewTempFolder() =>
        Path.Combine(Path.GetTempPath(), "FormFiller_시험", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
}
