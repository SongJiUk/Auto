using System.IO;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Microsoft.Win32;

namespace FormFiller.App.ViewModels;

/// <summary>
/// 메인 화면 사진 섹션에서 하나의 사진 슬롯을 나타내는 ViewModel.
///
/// [슬롯이란?]
/// 사진대지 시트에 미리 위치가 잡혀 있는 이미지 자리.
/// 서식 파일에서 anchor 개수를 세어 자동으로 생성된다.
///
/// 사용자가 사진을 선택하거나 드래그&드롭하면 SelectedImagePath가 채워지고,
/// 생성 시 PhotoService.ResizeToFill + ExcelPhotoWriter.ReplacePhoto로 삽입된다.
/// </summary>
public class PhotoSlotViewModel : ViewModelBase
{
    // ── 불변 메타 ──────────────────────────────────────────────────────

    /// <summary>문서 ID (어느 문서의 사진인지)</summary>
    public string DocumentId { get; }

    /// <summary>사진을 넣을 시트 이름</summary>
    public string SheetName { get; }

    /// <summary>이 슬롯의 순서 (0-based)</summary>
    public int SlotIndex { get; }

    /// <summary>표시 이름 (예: "사진 1")</summary>
    public string Label { get; }

    /// <summary>슬롯의 가득채우기 목표 너비(px)</summary>
    public int TargetWidthPx { get; }

    /// <summary>슬롯의 가득채우기 목표 높이(px)</summary>
    public int TargetHeightPx { get; }

    // ── 상태 프로퍼티 ───────────────────────────────────────────────

    private string? _selectedImagePath;

    /// <summary>
    /// 선택된 사진 파일 경로.
    /// null이면 미선택 상태 (빈 슬롯).
    /// </summary>
    public string? SelectedImagePath
    {
        get => _selectedImagePath;
        set
        {
            if (SetField(ref _selectedImagePath, value))
            {
                OnPropertyChanged(nameof(HasPhoto));
                OnPropertyChanged(nameof(PlaceholderVisible));
                LoadThumbnail(value);
            }
        }
    }

    /// <summary>사진이 선택되어 있으면 true</summary>
    public bool HasPhoto        => _selectedImagePath != null;

    /// <summary>사진이 없으면 플레이스홀더 표시</summary>
    public bool PlaceholderVisible => !HasPhoto;

    private BitmapImage? _thumbnail;

    /// <summary>미리보기용 BitmapImage (Image 컨트롤에 바인딩)</summary>
    public BitmapImage? Thumbnail
    {
        get => _thumbnail;
        private set => SetField(ref _thumbnail, value);
    }

    // ── 커맨드 ────────────────────────────────────────────────────────

    /// <summary>파일 대화상자로 사진을 선택한다</summary>
    public ICommand BrowseCommand { get; }

    /// <summary>선택된 사진을 지운다</summary>
    public ICommand ClearCommand { get; }

    // ── 생성자 ───────────────────────────────────────────────────────

    public PhotoSlotViewModel(
        string documentId, string sheetName, int slotIndex,
        int targetWidthPx, int targetHeightPx, string? label = null)
    {
        DocumentId     = documentId;
        SheetName      = sheetName;
        SlotIndex      = slotIndex;
        Label          = !string.IsNullOrWhiteSpace(label) ? label : $"사진 {slotIndex + 1}";
        TargetWidthPx  = targetWidthPx > 0 ? targetWidthPx : 668;
        TargetHeightPx = targetHeightPx > 0 ? targetHeightPx : 410;

        BrowseCommand = new RelayCommand(_ => OnBrowse());
        ClearCommand  = new RelayCommand(_ => SelectedImagePath = null, _ => HasPhoto);
    }

    // ── 메서드 ───────────────────────────────────────────────────────

    private void OnBrowse()
    {
        var dlg = new OpenFileDialog
        {
            Title  = $"{Label} 선택",
            Filter = "이미지 파일 (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp",
        };
        if (dlg.ShowDialog() == true)
            SelectedImagePath = dlg.FileName;
    }

    private void LoadThumbnail(string? path)
    {
        Thumbnail = path != null ? PhotoService.LoadThumbnail(path) : null;
    }

    /// <summary>드래그&드롭으로 받은 파일 경로를 처리한다.</summary>
    public void HandleDroppedFile(string filePath)
    {
        string ext = Path.GetExtension(filePath).ToLowerInvariant();
        if (ext is ".jpg" or ".jpeg" or ".png" or ".bmp")
            SelectedImagePath = filePath;
    }
}
