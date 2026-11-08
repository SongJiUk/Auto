using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using FormFiller.Core.Config;

namespace FormFiller.App.Views;

/// <summary>
/// 작업사항 항목 선택 팝업.
///
/// [동작]
/// - config.json 의 PresetWorkItems 목록을 체크박스로 표시
/// - 체크 후 "선택 추가" 클릭 → SelectedItems 에 담아 반환
/// - "목록에 추가": 직접 입력한 텍스트를 preset 목록에 영구 저장
/// - "바로 추가": preset 저장 없이 이번 선택에만 포함
/// </summary>
public partial class WorkItemPickerWindow : Window
{
    private readonly List<CheckBox> _checkBoxes = [];
    private AppConfig _config;

    /// <summary>확인 후 선택된 항목들. DialogResult=true 일 때만 유효.</summary>
    public List<string> SelectedItems { get; } = [];

    public WorkItemPickerWindow()
    {
        InitializeComponent();
        _config = AppConfigLoader.Load();
        BuildPresetList();
        UpdateHint();
    }

    // ── 체크박스 목록 구성 ────────────────────────────────────────────

    private void BuildPresetList()
    {
        PresetPanel.Children.Clear();
        _checkBoxes.Clear();

        foreach (var item in _config.PresetWorkItems)
        {
            var cb = new CheckBox
            {
                Content = item,
                Margin  = new Thickness(4, 4, 4, 2),
                Padding = new Thickness(4, 2, 0, 2),
                FontSize = 13,
            };
            cb.Checked   += (_, _) => UpdateHint();
            cb.Unchecked += (_, _) => UpdateHint();

            // 우클릭 → 목록에서 삭제
            cb.MouseRightButtonUp += (s, _) =>
            {
                if (s is not CheckBox target) return;
                var text = target.Content?.ToString() ?? "";
                _config.PresetWorkItems.Remove(text);
                AppConfigLoader.Save(_config);
                BuildPresetList();
            };
            cb.ToolTip = "우클릭: 목록에서 삭제";

            _checkBoxes.Add(cb);
            PresetPanel.Children.Add(cb);
        }
    }

    private void UpdateHint()
    {
        int count = _checkBoxes.Count(c => c.IsChecked == true);
        HintText.Text = count > 0 ? $"{count}개 선택됨" : "";
    }

    // ── 이벤트 핸들러 ─────────────────────────────────────────────────

    /// <summary>직접 입력 → 목록에 영구 저장 후 체크박스 재구성</summary>
    private void AddToPreset_Click(object sender, RoutedEventArgs e)
    {
        string text = CustomTextBox.Text.Trim();
        if (string.IsNullOrEmpty(text)) return;

        if (!_config.PresetWorkItems.Contains(text))
        {
            _config.PresetWorkItems.Add(text);
            AppConfigLoader.Save(_config);
            BuildPresetList();
            // 방금 추가한 항목 체크
            var last = _checkBoxes.LastOrDefault();
            if (last != null) last.IsChecked = true;
        }

        CustomTextBox.Text = "";
        CustomTextBox.Focus();
    }

    /// <summary>직접 입력 → 목록 저장 없이 이번 선택에만 추가</summary>
    private void AddDirect_Click(object sender, RoutedEventArgs e)
    {
        string text = CustomTextBox.Text.Trim();
        if (string.IsNullOrEmpty(text)) return;

        SelectedItems.Add(text);
        CustomTextBox.Text = "";
        DialogResult = true;
    }

    /// <summary>엔터 키: 바로 추가와 동일</summary>
    private void CustomTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter) AddDirect_Click(sender, new RoutedEventArgs());
    }

    /// <summary>체크된 항목 수집 후 창 닫기</summary>
    private void Confirm_Click(object sender, RoutedEventArgs e)
    {
        foreach (var cb in _checkBoxes.Where(c => c.IsChecked == true))
            SelectedItems.Add(cb.Content?.ToString() ?? "");

        DialogResult = SelectedItems.Count > 0;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
