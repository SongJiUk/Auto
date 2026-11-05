using System.IO;
using System.Windows.Media.Imaging;
using FormFiller.Core.Config;

namespace FormFiller.App.ViewModels.Settings;

/// <summary>
/// 설정 화면에서 사진 슬롯 하나를 나타내는 편집 행.
/// 슬롯의 플레이스홀더 이미지 썸네일을 보여주고, 사용자가 이름을 붙일 수 있다.
/// </summary>
public class PhotoSlotLabelRow : ViewModelBase
{
    /// <summary>시트 내 슬롯 순서 (0-based)</summary>
    public int SlotIndex { get; }

    /// <summary>슬롯 크기 힌트 (예: "668 × 410 px")</summary>
    public string SizeHint { get; }

    /// <summary>
    /// 서식 파일에 들어 있는 플레이스홀더 이미지 썸네일.
    /// null이면 "미리보기 없음" 표시.
    /// </summary>
    public BitmapImage? Thumbnail { get; }

    private string _label;

    /// <summary>사용자 지정 이름. 메인 화면 사진 카드 제목으로 표시된다.</summary>
    public string Label
    {
        get => _label;
        set => SetField(ref _label, value);
    }

    public PhotoSlotLabelRow(int slotIndex, int widthPx, int heightPx, string existingLabel, byte[]? imageBytes)
    {
        SlotIndex = slotIndex;
        SizeHint  = $"{widthPx} × {heightPx} px";
        _label    = existingLabel;
        Thumbnail = imageBytes != null ? LoadBitmap(imageBytes) : null;
    }

    public PhotoSlotConfig ToConfig() => new() { Index = SlotIndex, Label = Label };

    private static BitmapImage? LoadBitmap(byte[] bytes)
    {
        try
        {
            var bmp = new BitmapImage();
            bmp.BeginInit();
            bmp.StreamSource       = new MemoryStream(bytes);
            bmp.DecodePixelWidth   = 120;   // 썸네일 너비 (메모리 절약)
            bmp.CacheOption        = BitmapCacheOption.OnLoad;
            bmp.EndInit();
            bmp.Freeze();
            return bmp;
        }
        catch { return null; }
    }
}
