using System.IO;
using FormFiller.Core.Config;

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
        if (!TemplateMissing) Warning = $"저장 파일 이름이 '{others}'와 같아 서로 덮어씁니다 — 📁 문서 관리에서 바꿔 주세요";
    }

    /// <summary>오늘 생성해야 하는지 여부 (주기 계산 결과)</summary>
    public bool IsDue { get; }

    /// <summary>매일 업무인지 (아니면 매주·매월 등 주기 업무 묶음에 보인다)</summary>
    public bool IsDaily { get; }

    /// <summary>다음 생성 예정일 텍스트 (IsDue=false 일 때 표시, IsDue=true 이면 빈 문자열)</summary>
    public string NextDueText { get; }

    /// <summary>주기 업무의 이번 주기 상태 (예: "이번 주 완료 (11/16)"). 매일 업무는 빈 문자열.</summary>
    public string StatusText { get; }

    /// <summary>이번 주기에 이미 만들었는지 (초록 ✔ 표시)</summary>
    public bool IsDoneThisPeriod { get; }

    public DocumentDueItem(DocumentConfig config, DateTime today, string outputFolder, string baseDir)
    {
        Config = config;
        today  = today.Date;

        // 주기 계산은 Core(DocumentSchedule)에서 한다. 고른 날짜 이후 파일은 보지 않는다 (지난 날짜를 다시 만들 때)
        DateTime? last = DocumentSchedule.LastGenerated(config, Path.Combine(baseDir, outputFolder), notAfter: today);
        IsDue      = DocumentSchedule.IsDue(config, today, last);
        IsDaily    = DocumentSchedule.IsDaily(config);
        StatusText = DocumentSchedule.StatusText(config, today, last);
        IsDoneThisPeriod = !IsDaily && !IsDue && last != null;

        // 오늘 해당하면 기본 체크 상태로 시작
        // 서식 파일이 없으면 만들 수 없으니 처음부터 체크하지 않고 이유를 보여 준다
        TemplateMissing = string.IsNullOrWhiteSpace(config.TemplateFile)
                          || !File.Exists(AppConfigLoader.ResolvePath(config.TemplateFile));
        if (TemplateMissing) _warning = "서식 파일이 없습니다 — 📁 문서 관리에서 다시 골라 주세요";
        else if (config.HasRemovedInputFields())
            _warning = "예전 방식 칸이 있어 일부가 비어서 만들어집니다 — 📁 문서 관리 → ✏ 칸 편집에서 다시 지정해 주세요";
        _isChecked = IsDue && !TemplateMissing;

        FrequencyLabel = DocumentSchedule.Label(config);
        NextDueText = IsDue || !IsDaily ? "" : $"다음: {DocumentSchedule.NextDue(config, today, last):MM'/'dd}";
    }
}
