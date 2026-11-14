using System.IO;
using FormFiller.Core.Config;
using FormFiller.Core.Excel;

namespace FormFiller.App.ViewModels;

/// <summary>
/// 메인화면 "오늘 생성할 문서" 목록의 항목 하나.
/// 문서 설정과 오늘 생성 여부(체크박스)를 담는다.
/// </summary>
public class DocumentDueItem : ViewModelBase
{
    public DocumentConfig Config { get; }

    /// <summary>문서 이름 (config.Name 에서 그대로 가져온다)</summary>
    public string Name => Config.Name;

    /// <summary>주기 레이블 — UI 표시용 ("매일", "2주마다", "매월" 등)</summary>
    public string FrequencyLabel { get; }

    private bool _isChecked;

    /// <summary>이 문서를 오늘 생성할지 여부. 기본값은 IsDue 값과 동일하다.</summary>
    public bool IsChecked
    {
        get => _isChecked;
        set => SetField(ref _isChecked, value);
    }

    public bool TemplateMissing { get; }

    private string _warning = "";
    /// <summary>문서 이름 아래 빨간 글씨 (서식 없음, 저장 이름 겹침). 비어 있으면 안 보인다.</summary>
    public string Warning
    {
        get => _warning;
        private set { if (SetField(ref _warning, value)) OnPropertyChanged(nameof(HasWarning)); }
    }
    public bool HasWarning => _warning.Length > 0;

    /// <summary>다른 문서와 저장 파일 이름이 같을 때 (서식 없음 경고가 먼저)</summary>
    public void SetDuplicateName(string others)
    {
        if (!TemplateMissing) Warning = $"저장 파일 이름이 '{others}'와 같아 서로 덮어씁니다 — 설정에서 바꿔 주세요";
    }

    /// <summary>오늘 생성해야 하는지 여부 (주기 계산 결과)</summary>
    public bool IsDue { get; }

    /// <summary>다음 생성 예정일 텍스트 (IsDue=false 일 때 표시, IsDue=true 이면 빈 문자열)</summary>
    public string NextDueText { get; }

    public DocumentDueItem(DocumentConfig config, DateTime today, string outputFolder, string baseDir)
    {
        Config = config;
        IsDue  = CalculateIsDue(config, today, outputFolder, baseDir);

        // 오늘 해당하면 기본 체크 상태로 시작
        // 서식 파일이 없으면 만들 수 없으니 처음부터 체크하지 않고 이유를 보여 준다
        TemplateMissing = string.IsNullOrWhiteSpace(config.TemplateFile)
                          || !File.Exists(AppConfigLoader.ResolvePath(config.TemplateFile));
        if (TemplateMissing) _warning = "서식 파일이 없습니다 — ⚙ 설정에서 다시 골라 주세요";
        _isChecked = IsDue && !TemplateMissing;

        FrequencyLabel = config.Frequency switch
        {
            "daily"     => "매일",
            "biweekly"  => "2주마다",
            "monthly"   => "매월",
            "custom"    => $"{config.FrequencyDays}일마다",
            _           => ""
        };

        NextDueText = IsDue
            ? ""
            : $"다음: {CalculateNextDue(config, today, outputFolder, baseDir):MM'/'dd}";
    }

    // ── 주기 계산 ─────────────────────────────────────────────────────

    /// <summary>
    /// 오늘 이 문서를 생성해야 하는지 판단한다.
    /// output 폴더에서 가장 최근 파일을 찾아 마지막 생성일을 추정한다.
    /// </summary>
    private static bool CalculateIsDue(DocumentConfig config, DateTime today,
                                        string outputFolder, string baseDir)
    {
        // daily 는 항상 생성
        if (config.Frequency == "daily") return true;

        DateTime? lastGenerated = FindLastGeneratedDate(config, outputFolder, baseDir);

        // 한 번도 생성하지 않았으면 오늘 생성
        if (lastGenerated == null) return true;

        int daysSinceLast = (today.Date - lastGenerated.Value.Date).Days;

        return config.Frequency switch
        {
            "biweekly" => daysSinceLast >= 14,
            "monthly"  => today.Month != lastGenerated.Value.Month
                          || today.Year != lastGenerated.Value.Year,
            "custom"   => daysSinceLast >= config.FrequencyDays,
            _          => true
        };
    }

    /// <summary>다음 생성 예정일을 계산한다.</summary>
    private static DateTime CalculateNextDue(DocumentConfig config, DateTime today,
                                              string outputFolder, string baseDir)
    {
        DateTime last = FindLastGeneratedDate(config, outputFolder, baseDir) ?? today;

        return config.Frequency switch
        {
            "biweekly" => last.AddDays(14),
            "monthly"  => new DateTime(last.Year, last.Month, 1).AddMonths(1),
            "custom"   => last.AddDays(config.FrequencyDays),
            _          => today
        };
    }

    // ── 파일 날짜 추정 ────────────────────────────────────────────────

    /// <summary>
    /// output 폴더에서 이 문서의 패턴에 맞는 파일 중 가장 최근 날짜를 찾는다.
    /// OutputPattern 의 고정 prefix 를 추출한 뒤 파일명 뒤의 yy.MM.dd 를 파싱한다.
    /// </summary>
    private static DateTime? FindLastGeneratedDate(DocumentConfig config,
                                                    string outputFolder, string baseDir)
    {
        string dir = Path.Combine(baseDir, outputFolder);
        if (!Directory.Exists(dir)) return null;

        // 패턴에서 고정 prefix 추출 (예: "공사일보_출력일보_" ← "{yy}" 앞 부분)
        string prefix = config.OutputPattern.Split('{')[0];
        DateTime? latest = null;

        foreach (string file in Directory.EnumerateFiles(dir, "*.xlsx"))
        {
            string fileName = Path.GetFileNameWithoutExtension(file);
            if (!fileName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;

            // prefix 이후 부분에서 yy.MM.dd 파싱 시도
            string datePart = fileName[prefix.Length..];
            if (TryParseFileDate(datePart, out DateTime parsed))
            {
                if (latest == null || parsed > latest)
                    latest = parsed;
            }
        }

        return latest;
    }

    /// <summary>
    /// "26.11.04" → 2026-11-04 형태로 파싱한다.
    /// 파싱 실패 시 false 를 반환한다.
    /// </summary>
    private static bool TryParseFileDate(string s, out DateTime result)
    {
        if (DateTime.TryParseExact(
                s, "yy.MM.dd",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None,
                out result))
        {
            return true;
        }

        result = default;
        return false;
    }
}
