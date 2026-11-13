using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows.Media.Imaging;

namespace FormFiller.App;

/// <summary>
/// 측정기(소음계) 사진에서 화면에 표시된 숫자를 Claude API 로 읽는다.
///
/// [원칙] 결과는 "제안"일 뿐이다. 입력 칸에 채워 주기만 하고, 사람이 확인한 뒤 생성한다.
/// 사용자가 [읽기] 버튼을 누른 사진 1장만 전송한다 (작게 줄여서 보낸다).
/// </summary>
public static class MeterReader
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(60) };
    private const string ApiUrl = "https://api.anthropic.com/v1/messages";
    private const string Model  = "claude-haiku-5-5";

    /// <summary>사진 속 측정기 숫자를 읽어 "55.3" 같은 글자로 돌려준다. 못 읽으면 예외.</summary>
    public static async Task<string> ReadAsync(string apiKey, string imagePath, string? unit)
    {
        if (string.IsNullOrWhiteSpace(apiKey) || apiKey.Trim().Length < 20)
            throw new InvalidOperationException(
                "Claude API 키가 없습니다.\n메인 화면의 AI 도우미에서 API 키를 먼저 저장해 주세요.");

        // 긴 변 1200px 로 줄인 JPEG (회전 보정 포함) → base64
        var src = PhotoService.LoadOriented(imagePath, 1200);
        var enc = new JpegBitmapEncoder { QualityLevel = 85 };
        enc.Frames.Add(BitmapFrame.Create(src));
        using var ms = new MemoryStream();
        enc.Save(ms);
        string b64 = Convert.ToBase64String(ms.ToArray());

        string prompt =
            "이 사진은 공사 현장의 측정기(예: 소음계) 화면입니다. 화면에 크게 표시된 측정값 숫자만 답하세요"
            + (string.IsNullOrEmpty(unit) ? "" : $" (단위 {unit} 는 빼고)")
            + ". 예: 55.3\n숫자를 확실히 읽을 수 없으면 ? 한 글자만 답하세요.";

        var body = new
        {
            model = Model,
            max_tokens = 20,
            messages = new object[]
            {
                new
                {
                    role = "user",
                    content = new object[]
                    {
                        new { type = "image", source = new { type = "base64", media_type = "image/jpeg", data = b64 } },
                        new { type = "text", text = prompt },
                    },
                },
            },
        };

        using var req = new HttpRequestMessage(HttpMethod.Post, ApiUrl)
        {
            Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json"),
        };
        req.Headers.Add("x-api-key", apiKey.Trim());
        req.Headers.Add("anthropic-version", "2023-06-01");

        using var res = await Http.SendAsync(req);
        string json = await res.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(json);

        if (!res.IsSuccessStatusCode)
        {
            string msg = doc.RootElement.TryGetProperty("error", out var err)
                         && err.TryGetProperty("message", out var m) ? m.GetString() ?? json : json;
            throw new InvalidOperationException($"API 오류: {msg}");
        }

        string text = doc.RootElement.GetProperty("content")[0].GetProperty("text").GetString() ?? "";
        var match = Regex.Match(text, @"\d+(\.\d+)?");
        if (!match.Success)
            throw new InvalidOperationException("사진에서 숫자를 읽지 못했습니다. 직접 입력해 주세요.");
        return match.Value;
    }
}
