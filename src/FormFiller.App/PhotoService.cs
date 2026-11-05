using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace FormFiller.App;

/// <summary>
/// WPF 이미징을 사용해 사진을 처리하는 서비스.
///
/// [역할]
/// - ResizeToFill : 사진을 지정 픽셀 크기로 늘려(가득채우기) JPEG 바이트 배열 반환
/// - LoadThumbnail: 미리보기용 BitmapImage 생성 (UI 스레드에서 호출)
/// </summary>
public static class PhotoService
{
    /// <summary>
    /// 이미지 파일을 targetW × targetH 픽셀로 "가득채우기(stretch)" 리사이즈해
    /// JPEG 바이너리로 반환한다.
    ///
    /// [동작]
    /// DrawingVisual에 원본 이미지를 targetW × targetH 사각형에 그린 뒤
    /// RenderTargetBitmap으로 래스터화하면 가로세로 비율과 무관하게 꽉 채워진다.
    /// (Excel drawing XML의 &lt;a:stretch&gt;와 동일한 효과)
    /// </summary>
    public static byte[] ResizeToFill(string imagePath, int targetW, int targetH)
    {
        // BitmapImage 로드 (파일을 미리 메모리에 올려 파일 잠금 방지)
        var bmp = new BitmapImage();
        bmp.BeginInit();
        bmp.UriSource    = new Uri(imagePath, UriKind.Absolute);
        bmp.CacheOption  = BitmapCacheOption.OnLoad;
        bmp.EndInit();
        bmp.Freeze();

        // DrawingVisual로 targetW × targetH 사각형에 이미지를 꽉 채워 그린다
        var visual = new DrawingVisual();
        using (var dc = visual.RenderOpen())
            dc.DrawImage(bmp, new Rect(0, 0, targetW, targetH));

        // 96 DPI 비트맵으로 래스터화
        var rtb = new RenderTargetBitmap(targetW, targetH, 96, 96, PixelFormats.Pbgra32);
        rtb.Render(visual);
        rtb.Freeze();

        // JPEG 인코딩 (품질 90)
        var encoder = new JpegBitmapEncoder { QualityLevel = 90 };
        encoder.Frames.Add(BitmapFrame.Create(rtb));

        using var ms = new MemoryStream();
        encoder.Save(ms);
        return ms.ToArray();
    }

    /// <summary>
    /// 미리보기용 BitmapImage를 로드한다 (최대 decodeWidth 픽셀로 축소).
    /// UI에 직접 바인딩할 수 있도록 Freeze()된 상태로 반환된다.
    /// </summary>
    public static BitmapImage? LoadThumbnail(string imagePath, int decodeWidth = 160)
    {
        if (!File.Exists(imagePath)) return null;

        try
        {
            var bmp = new BitmapImage();
            bmp.BeginInit();
            bmp.UriSource       = new Uri(imagePath, UriKind.Absolute);
            bmp.CacheOption     = BitmapCacheOption.OnLoad;
            bmp.DecodePixelWidth = decodeWidth; // 메모리 절약: 필요한 너비만큼만 디코딩
            bmp.EndInit();
            bmp.Freeze();
            return bmp;
        }
        catch
        {
            return null;
        }
    }
}
