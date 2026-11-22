using FormFiller.Core.Config;

namespace FormFiller.App.ViewModels.Settings;

/// <summary>
/// 문서 관리 화면에서 한 항목(FieldConfig)을 보여 주는 읽기 전용 행.
/// 칸·규칙을 바꾸는 곳은 [✏ 칸 편집] 하나뿐이라 여기서는 이름·종류·칸 요약만 보여 준다 (2026-11-22 화면 정리).
/// </summary>
public class FieldEditRow
{
    private readonly FieldConfig _orig;

    public FieldEditRow(FieldConfig f) => _orig = f;

    public string Label => _orig.Label;

    /// <summary>"값", "표" 같은 종류 이름. 예전 방식 항목은 [칸 편집]에서 다시 지정하라는 뜻으로 따로 표시.</summary>
    public string TypeTitle => BlockTypes.IsBlock(_orig.Type) ? BlockTypes.TypeTitle(_orig.Type) : "예전 방식";

    public bool IsLegacy => !BlockTypes.IsBlock(_orig.Type);

    public string CellSummary => BlockTypes.CellSummary(_orig);

    /// <summary>값 이름(Key). 다른 문서와 같은 값을 나눠 쓰면 여기 이름이 같다.</summary>
    public string KeyText => string.IsNullOrWhiteSpace(_orig.Key) || _orig.Key.Trim() == _orig.Label.Trim()
                             ? "" : _orig.Key.Trim();

    /// <summary>이 화면에서는 고치지 않으므로 원본 그대로 (복사본)</summary>
    public FieldConfig ToFieldConfig() =>
        System.Text.Json.JsonSerializer.Deserialize<FieldConfig>(System.Text.Json.JsonSerializer.Serialize(_orig))!;
}
