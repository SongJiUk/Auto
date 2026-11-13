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

    // ── 폰 사진용 (회전 보정 + 잘라서 채우기 + 촬영 시각) ──────────────

    /// <summary>
    /// 사진을 targetW × targetH 칸에 "비율 유지하며 꽉 채우고 넘치는 부분은 가운데 기준으로 잘라" JPEG 로 만든다.
    /// 측정기 화면 사진처럼 찌그러지면 안 되는 사진에 쓴다.
    /// 폰 사진의 회전 정보(EXIF Orientation)를 반영해 눕지 않게 한다.
    /// </summary>
    public static byte[] ResizeToCover(string imagePath, int targetW, int targetH)
    {
        var src = LoadOriented(imagePath, Math.Max(targetW, targetH) * 2);

        double scale = Math.Max(targetW / src.Width, targetH / src.Height);
        double w = src.Width * scale, h = src.Height * scale;

        var visual = new DrawingVisual();
        using (var dc = visual.RenderOpen())
            dc.DrawImage(src, new Rect((targetW - w) / 2, (targetH - h) / 2, w, h));

        var rtb = new RenderTargetBitmap(targetW, targetH, 96, 96, PixelFormats.Pbgra32);
        rtb.Render(visual);

        var encoder = new JpegBitmapEncoder { QualityLevel = 90 };
        encoder.Frames.Add(BitmapFrame.Create(rtb));
        using var ms = new MemoryStream();
        encoder.Save(ms);
        return ms.ToArray();
    }

    /// <summary>
    /// 사진을 읽어 EXIF 회전 정보대로 돌린 이미지를 돌려준다.
    /// (폰은 사진을 눕힌 채 저장하고 "몇 도 돌려서 보라"는 표시만 남기는 경우가 많다)
    /// </summary>
    /// <param name="maxSide">긴 변 최대 픽셀 (메모리 절약용 축소, 0 이면 원본)</param>
    public static BitmapSource LoadOriented(string imagePath, int maxSide = 0)
    {
        byte[] data = File.ReadAllBytes(imagePath);
        int orientation = ReadOrientation(data);

        var bmp = new BitmapImage();
        bmp.BeginInit();
        bmp.StreamSource = new MemoryStream(data);
        bmp.CacheOption  = BitmapCacheOption.OnLoad;
        // 회전 전 기준으로 줄인다 (90도 회전 사진은 높이가 긴 변이 된다)
        if (maxSide > 0)
        {
            if (orientation is 6 or 8) bmp.DecodePixelHeight = maxSide;
            else                       bmp.DecodePixelWidth  = maxSide;
        }
        bmp.EndInit();
        bmp.Freeze();

        Transform? t = orientation switch
        {
            3 => new RotateTransform(180),
            6 => new RotateTransform(90),
            8 => new RotateTransform(270),
            _ => null,
        };
        if (t == null) return bmp;

        var tb = new TransformedBitmap(bmp, t);
        tb.Freeze();
        return tb;
    }

    /// <summary>회전 보정된 미리보기 (폰 사진 썸네일용)</summary>
    public static BitmapSource? LoadOrientedThumbnail(string imagePath, int maxSide = 240)
    {
        try { return File.Exists(imagePath) ? LoadOriented(imagePath, maxSide) : null; }
        catch { return null; }
    }

    /// <summary>
    /// 촬영 시각. 사진 정보(EXIF)에 없으면 파일 수정 시각을 쓴다.
    /// (구글 드라이브로 받은 사진도 보통 촬영 시각 정보가 남아 있다)
    /// </summary>
    public static DateTime GetTakenTime(string imagePath)
    {
        try
        {
            using var fs = File.OpenRead(imagePath);
            var frame = BitmapFrame.Create(fs, BitmapCreateOptions.DelayCreation, BitmapCacheOption.None);
            if (frame.Metadata is BitmapMetadata md && DateTime.TryParse(md.DateTaken, out var taken))
                return taken;
        }
        catch { /* 사진 정보를 못 읽으면 파일 시각 사용 */ }
        return File.GetLastWriteTime(imagePath);
    }

    private static int ReadOrientation(byte[] data)
    {
        try
        {
            var frame = BitmapFrame.Create(new MemoryStream(data), BitmapCreateOptions.DelayCreation, BitmapCacheOption.None);
            if (frame.Metadata is BitmapMetadata md
                && md.ContainsQuery("System.Photo.Orientation")
                && md.GetQuery("System.Photo.Orientation") is ushort o)
                return o;
        }
        catch { }
        return 1;
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
