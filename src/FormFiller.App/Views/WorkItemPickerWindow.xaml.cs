using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using FormFiller.Core.Config;

namespace FormFiller.App.Views;

/// <summary>
/// [☰ 항목 선택] 팝업. 여러 줄 항목(값 이름)마다 목록이 따로 있다 (작업사항·장비 이름·장비 규격…).
///
/// [동작]
/// - config.json 의 ItemPickLists[항목 이름] 을 체크박스로 표시 (규칙은 Core PickLists)
/// - 지난번에 고른 것은 미리 체크돼 있다 → 다음 날 켜도 그대로
/// - "목록에 추가": 새 항목을 목록에 영구 저장 / "바로 추가": 이번에만
/// - 줄 오른쪽 ✕: 목록에서 삭제
/// </summary>
public partial class WorkItemPickerWindow : Window
{
    private readonly List<CheckBox> _checkBoxes = [];
    private readonly AppConfig _config;
    private readonly string _key;

    /// <summary>확인 후 선택된 항목들. DialogResult=true 일 때만 유효.</summary>
    public List<string> SelectedItems { get; } = [];

    /// <param name="key">항목 이름 (예: "금일 작업사항"). 이름마다 목록이 따로 저장된다.</param>
    public WorkItemPickerWindow(string key)
    {
        InitializeComponent();
        _key    = key;
        _config = AppConfigLoader.Load();
        Title   = $"{key} — 항목 선택";
        HeaderText.Text = $"'{key}' 에 넣을 것을 체크하세요. 목록은 이 항목에만 저장되고, 지난번에 고른 것은 미리 체크됩니다.";
        BuildPresetList(PickLists.LastChosen(_config, _key));
        UpdateHint();
    }

    // ── 체크박스 목록 구성 ────────────────────────────────────────────

    /// <param name="checkedItems">체크해 둘 항목</param>
    private void BuildPresetList(IEnumerable<string> checkedItems)
    {
        var toCheck = new HashSet<string>(checkedItems);
        PresetPanel.Children.Clear();
        _checkBoxes.Clear();

        var items = PickLists.Items(_config, _key);
        if (items.Count == 0)
        {
            PresetPanel.Children.Add(new TextBlock
            {
                Text = "아직 목록이 비어 있습니다. 아래에 쓰고 [목록에 추가] 를 누르세요.",
                Foreground = System.Windows.Media.Brushes.Gray, Margin = new Thickness(6), TextWrapping = TextWrapping.Wrap,
            });
            return;
        }

        foreach (var item in items)
        {
            var cb = new CheckBox
            {
                Content   = item,
                IsChecked = toCheck.Contains(item),
                Margin    = new Thickness(4, 4, 4, 2),
                Padding   = new Thickness(4, 2, 0, 2),
                FontSize  = 13,
                VerticalAlignment = VerticalAlignment.Center,
            };
            cb.Checked   += (_, _) => UpdateHint();
            cb.Unchecked += (_, _) => UpdateHint();

            var del = new Button
            {
                Content = "✕", ToolTip = "목록에서 삭제", Padding = new Thickness(6, 0, 6, 0),
                Background = System.Windows.Media.Brushes.Transparent, BorderThickness = new Thickness(0),
                Foreground = System.Windows.Media.Brushes.IndianRed,
            };
            string text = item;
            del.Click += (_, _) => RemoveItem(text);

            var row = new DockPanel();
            DockPanel.SetDock(del, Dock.Right);
            row.Children.Add(del);
            row.Children.Add(cb);

            _checkBoxes.Add(cb);
            PresetPanel.Children.Add(row);
        }
    }

    private List<string> CheckedTexts() =>
        _checkBoxes.Where(c => c.IsChecked == true).Select(c => c.Content?.ToString() ?? "").ToList();

    private void RemoveItem(string text)
    {
        var keep = CheckedTexts();
        PickLists.Remove(_config, _key, text);
        AppConfigLoader.Save(_config);
        BuildPresetList(keep);
        UpdateHint();
    }

    private void UpdateHint()
    {
        int count = _checkBoxes.Count(c => c.IsChecked == true);
        HintText.Text = count > 0 ? $"{count}개 선택됨" : "";
    }

    // ── 이벤트 핸들러 ─────────────────────────────────────────────────

    /// <summary>새 항목 → 목록에 영구 저장 후 체크된 채로 다시 그림</summary>
    private void AddToPreset_Click(object sender, RoutedEventArgs e)
    {
        string text = CustomTextBox.Text.Trim();
        if (string.IsNullOrEmpty(text)) return;

        var keep = CheckedTexts();
        keep.Add(text);
        if (PickLists.Add(_config, _key, text))
            AppConfigLoader.Save(_config);
        BuildPresetList(keep);
        UpdateHint();

        CustomTextBox.Text = "";
        CustomTextBox.Focus();
    }

    /// <summary>직접 입력 → 목록 저장 없이 이번 선택에만 추가 (체크한 것도 같이 넣는다)</summary>
    private void AddDirect_Click(object sender, RoutedEventArgs e)
    {
        string text = CustomTextBox.Text.Trim();
        if (string.IsNullOrEmpty(text)) return;

        Confirm(extra: text);
    }

    /// <summary>엔터 키: 바로 추가와 동일</summary>
    private void CustomTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter) AddDirect_Click(sender, new RoutedEventArgs());
    }

    /// <summary>체크된 항목 수집 후 창 닫기</summary>
    private void Confirm_Click(object sender, RoutedEventArgs e) => Confirm(extra: null);

    private void Confirm(string? extra)
    {
        var chosen = CheckedTexts();
        SelectedItems.AddRange(chosen);
        if (extra != null) SelectedItems.Add(extra);

        PickLists.SetLastChosen(_config, _key, chosen);   // 다음에 열면 이것들이 체크돼 있다
        AppConfigLoader.Save(_config);

        DialogResult = SelectedItems.Count > 0;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
