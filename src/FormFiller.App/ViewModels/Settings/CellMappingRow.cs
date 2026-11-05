using FormFiller.Core.Excel;

namespace FormFiller.App.ViewModels.Settings;

/// <summary>
/// 설정 화면에서 "항목명 | 셀 주소 입력 | 미리보기" 한 행을 나타내는 ViewModel.
///
/// [역할]
/// - Label     : "날짜", "날씨" 등 — 사용자에게 보이는 항목 이름 (읽기 전용)
/// - FieldId   : "DateCell" 등 — config 프로퍼티명과 매핑할 때 쓰는 식별자
/// - CellAddress : "C6" 등 — 사용자가 직접 편집, 바뀌면 미리보기를 자동 갱신
/// - Preview   : 실제 엑셀 파일에서 읽어 온 현재 값 (읽기 전용, 힌트용)
/// </summary>
public class CellMappingRow : ViewModelBase
{
    // ── 읽기 전용 식별 정보 ───────────────────────────────────────────

    /// <summary>화면에 표시할 항목 이름 ("날짜", "날씨" 등)</summary>
    public string Label { get; }

    /// <summary>config 프로퍼티와 매핑하는 식별자 ("DateCell", "WeatherCell" 등)</summary>
    public string FieldId { get; }

    // ── 외부에서 주입받는 파일·시트 정보 (파일 선택·시트 변경 시 갱신) ─

    // 미리보기를 가져올 파일 경로 (SettingsViewModel 이 주입)
    private string _filePath = "";

    // 미리보기를 가져올 시트 이름 (SettingsViewModel 이 주입)
    private string _sheetName = "";

    // ── 사용자 편집 가능 프로퍼티 ────────────────────────────────────

    private string _cellAddress;

    /// <summary>
    /// 사용자가 입력하는 셀 주소 (예: "C6").
    /// 값이 바뀌면 미리보기를 자동으로 갱신한다.
    /// </summary>
    public string CellAddress
    {
        get => _cellAddress;
        set
        {
            if (SetField(ref _cellAddress, value))
                RefreshPreview(_filePath, _sheetName);
        }
    }

    // ── 읽기 전용 미리보기 ───────────────────────────────────────────

    private string _preview = "";

    /// <summary>
    /// 현재 엑셀 파일에서 읽어 온 셀 값 (힌트용, 읽기 전용).
    /// 파일을 선택하거나 시트가 바뀔 때 SettingsViewModel 이 RefreshPreview() 를 호출한다.
    /// </summary>
    public string Preview
    {
        get => _preview;
        private set => SetField(ref _preview, value);
    }

    // ── 생성자 ───────────────────────────────────────────────────────

    /// <summary>
    /// CellMappingRow 를 초기화한다.
    /// </summary>
    /// <param name="label">화면 항목명 ("날짜")</param>
    /// <param name="fieldId">config 프로퍼티 식별자 ("DateCell")</param>
    /// <param name="initialCellAddress">config 에서 읽어 온 초기 셀 주소 ("C6")</param>
    public CellMappingRow(string label, string fieldId, string initialCellAddress)
    {
        Label         = label;
        FieldId       = fieldId;
        _cellAddress  = initialCellAddress; // SetField 없이 초기값 세팅 (PropertyChanged 불필요)
    }

    // ── 미리보기 갱신 ────────────────────────────────────────────────

    /// <summary>
    /// 현재 셀 주소로 엑셀 파일에서 값을 읽어 Preview 를 갱신한다.
    ///
    /// filePath 나 sheetName 이 비어 있으면 Preview 를 비운다.
    /// 외부에서 파일/시트가 바뀔 때마다 SettingsViewModel 이 이 메서드를 호출한다.
    /// </summary>
    /// <param name="filePath">읽을 엑셀 파일 경로 (비어 있으면 미리보기 없음)</param>
    /// <param name="sheetName">읽을 시트 이름 (비어 있으면 미리보기 없음)</param>
    public void RefreshPreview(string filePath = "", string sheetName = "")
    {
        // 외부에서 주입한 파일·시트 정보를 업데이트
        _filePath  = filePath;
        _sheetName = sheetName;

        if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(sheetName))
        {
            Preview = "";
            return;
        }

        // ExcelAnalyzer 로 셀 값을 읽어 온다 (오류 시 "(읽기 오류)" 반환)
        string raw = ExcelAnalyzer.GetCellPreview(filePath, sheetName, CellAddress);

        // 긴 텍스트는 30자로 잘라 힌트처럼 보여 준다
        Preview = raw.Length > 30 ? raw[..30] + "…" : raw;
    }
}
