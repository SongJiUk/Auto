using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using FormFiller.Core.Config;

namespace FormFiller.App.Views;

/// <summary>
/// Claude AI 채팅 창.
/// Anthropic REST API 를 HttpClient 로 직접 호출한다 (외부 SDK 불필요).
/// API 키는 config.json 의 ClaudeApiKey 에 저장한다.
/// </summary>
public partial class AiAssistantWindow : Window
{
    // ── 상태 ──────────────────────────────────────────────────────────
    private static readonly HttpClient _http = new();
    private readonly List<ChatMessage> _history = [];
    private string _contextSummary = "";

    // Anthropic API 엔드포인트
    private const string ApiUrl   = "https://api.anthropic.com/v1/messages";
    private const string ApiModel = "claude-haiku-4-5-20251001";

    // ── 생성자 ────────────────────────────────────────────────────────

    public AiAssistantWindow()
    {
        InitializeComponent();
        RefreshApiKeyPanel();
        AddWelcomeMessage();
    }

    /// <summary>
    /// 현재 작업 현황(날짜·작업사항·장비·인원)을 AI 컨텍스트로 주입.
    /// MainWindow 에서 창을 열기 전에 호출한다.
    /// </summary>
    public void SetContext(string contextSummary)
    {
        _contextSummary = contextSummary;
    }

    // ── UI 초기화 ─────────────────────────────────────────────────────

    private void RefreshApiKeyPanel()
    {
        var config = AppConfigLoader.Load();
        ApiKeyPanel.Visibility = string.IsNullOrWhiteSpace(config.ClaudeApiKey)
            ? Visibility.Visible
            : Visibility.Collapsed;
        SendButton.IsEnabled = !string.IsNullOrWhiteSpace(config.ClaudeApiKey);
    }

    private void AddWelcomeMessage()
    {
        AddBubble("안녕하세요! 공사 서류 작성을 도와드릴게요.\n"
                + "예: \"오늘 작업 항목 정리해줘\", \"TBM 내용 추천해줘\"",
                  isUser: false);
    }

    // ── API 키 저장 ───────────────────────────────────────────────────

    private void SaveApiKey_Click(object sender, RoutedEventArgs e)
    {
        string key = ApiKeyBox.Text.Trim();
        if (string.IsNullOrEmpty(key)) return;

        var config = AppConfigLoader.Load();
        config.ClaudeApiKey = key;
        AppConfigLoader.Save(config);

        ApiKeyBox.Text = "";
        RefreshApiKeyPanel();
        AddBubble("API 키가 저장되었습니다. 이제 질문하세요!", isUser: false);
    }

    // ── 메시지 전송 ───────────────────────────────────────────────────

    private void InputBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && !e.IsRepeat)
            _ = SendMessageAsync();
    }

    private void Send_Click(object sender, RoutedEventArgs e)
        => _ = SendMessageAsync();

    private async Task SendMessageAsync()
    {
        string text = InputBox.Text.Trim();
        if (string.IsNullOrEmpty(text)) return;

        var config = AppConfigLoader.Load();
        if (string.IsNullOrWhiteSpace(config.ClaudeApiKey))
        {
            AddBubble("API 키를 먼저 입력해주세요.", isUser: false);
            return;
        }

        // 화면에 사용자 메시지 표시
        InputBox.Text = "";
        AddBubble(text, isUser: true);
        _history.Add(new ChatMessage("user", text));

        // 로딩 표시
        LoadingText.Visibility = Visibility.Visible;
        SendButton.IsEnabled   = false;
        InputBox.IsEnabled     = false;

        try
        {
            string reply = await CallClaudeAsync(config.ClaudeApiKey, _history);
            _history.Add(new ChatMessage("assistant", reply));
            AddBubble(reply, isUser: false);
        }
        catch (Exception ex)
        {
            AddBubble($"오류: {ex.Message}", isUser: false);
        }
        finally
        {
            LoadingText.Visibility = Visibility.Collapsed;
            SendButton.IsEnabled   = true;
            InputBox.IsEnabled     = true;
            InputBox.Focus();
        }
    }

    // ── Anthropic API 호출 ────────────────────────────────────────────

    private async Task<string> CallClaudeAsync(string apiKey, List<ChatMessage> history)
    {
        // 시스템 프롬프트: 역할 + 현재 현장 컨텍스트
        string systemPrompt =
            "당신은 한국 공사 현장의 서류 작성을 돕는 AI 도우미입니다. "
            + "공사일보, TBM일지, 사진대지 등 건설 현장 서류에 대해 잘 알고 있습니다. "
            + "한국어로 간결하고 실용적으로 답변하세요.\n\n"
            + (string.IsNullOrEmpty(_contextSummary)
                ? ""
                : $"현재 현장 현황:\n{_contextSummary}");

        // 메시지 배열 직렬화
        var messages = history.Select(m => new { role = m.Role, content = m.Content }).ToList();

        var requestBody = new
        {
            model      = ApiModel,
            max_tokens = 1024,
            system     = systemPrompt,
            messages,
        };

        string json    = JsonSerializer.Serialize(requestBody);
        var    content = new StringContent(json, Encoding.UTF8, "application/json");

        using var request = new HttpRequestMessage(HttpMethod.Post, ApiUrl);
        request.Headers.Add("x-api-key", apiKey);
        request.Headers.Add("anthropic-version", "2023-06-01");
        request.Content = content;

        var response = await _http.SendAsync(request);
        string responseBody = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            // 에러 메시지 파싱 시도
            try
            {
                using var errDoc = JsonDocument.Parse(responseBody);
                string errMsg = errDoc.RootElement
                    .GetProperty("error")
                    .GetProperty("message")
                    .GetString() ?? responseBody;
                throw new InvalidOperationException(errMsg);
            }
            catch (KeyNotFoundException) { throw new InvalidOperationException(responseBody); }
        }

        // 응답 파싱: content[0].text
        using var doc  = JsonDocument.Parse(responseBody);
        string    text = doc.RootElement
            .GetProperty("content")[0]
            .GetProperty("text")
            .GetString() ?? "(응답 없음)";

        return text;
    }

    // ── 채팅 버블 UI ──────────────────────────────────────────────────

    private void AddBubble(string text, bool isUser)
    {
        var border = new Border
        {
            CornerRadius    = new CornerRadius(10),
            Padding         = new Thickness(10, 7, 10, 7),
            Margin          = new Thickness(isUser ? 60 : 0, 3, isUser ? 0 : 60, 3),
            Background      = isUser
                                ? new SolidColorBrush(Color.FromRgb(43, 108, 176))
                                : new SolidColorBrush(Color.FromRgb(237, 242, 247)),
            HorizontalAlignment = isUser ? HorizontalAlignment.Right : HorizontalAlignment.Left,
        };

        var tb = new TextBlock
        {
            Text            = text,
            TextWrapping    = TextWrapping.Wrap,
            Foreground      = isUser
                                ? Brushes.White
                                : new SolidColorBrush(Color.FromRgb(45, 55, 72)),
            FontSize        = 13,
        };
        border.Child = tb;
        ChatPanel.Children.Add(border);

        // 스크롤 최하단
        ChatScroll.UpdateLayout();
        ChatScroll.ScrollToEnd();
    }

    // ── 내부 모델 ────────────────────────────────────────────────────

    private record ChatMessage(string Role, string Content);
}
