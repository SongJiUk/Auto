using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;

namespace FormFiller.App.Views;

/// <summary>
/// WPF HwndHost 를 상속해 Excel 창을 WPF 레이아웃 안에 직접 임베드한다.
/// WindowsFormsHost 를 쓰지 않으므로 Office COM 메시지 루프와 충돌이 없다.
///
/// [리사이즈] ArrangeOverride 에서 DPI 를 반영해 Win32 창 크기를 갱신한다.
///            전체화면·최대화 시에도 WPF 레이아웃 패스가 이 메서드를 호출하므로
///            Excel 창이 항상 올바른 크기를 유지한다.
/// </summary>
internal sealed class ExcelHwndHost : HwndHost
{
    // ── Win32 ────────────────────────────────────────────────────────
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    static extern IntPtr CreateWindowEx(int exStyle, string className, string? title,
        int style, int x, int y, int w, int h,
        IntPtr parent, IntPtr menu, IntPtr instance, IntPtr param);

    [DllImport("user32.dll")] static extern bool DestroyWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern IntPtr SetParent(IntPtr child, IntPtr parent);
    [DllImport("user32.dll")] static extern bool MoveWindow(IntPtr hwnd, int x, int y, int w, int h, bool repaint);
    [DllImport("user32.dll")] static extern int  GetWindowLong(IntPtr hwnd, int index);
    [DllImport("user32.dll")] static extern int  SetWindowLong(IntPtr hwnd, int index, int value);

    const int GWL_STYLE      = -16;
    const int WS_CHILD       = 0x40000000;
    const int WS_VISIBLE     = 0x10000000;
    const int WS_CLIPCHILDREN= 0x02000000;
    const int WS_CLIPSIBLINGS= 0x04000000;
    const int WS_CAPTION     = 0x00C00000;
    const int WS_THICKFRAME  = 0x00040000;
    const int WS_SYSMENU     = 0x00080000;
    const int WS_MAXIMIZEBOX = 0x00010000;
    const int WS_MINIMIZEBOX = 0x00020000;

    private IntPtr _containerHwnd = IntPtr.Zero;
    private IntPtr _excelHwnd     = IntPtr.Zero;

    // ── HwndHost 필수 구현 ──────────────────────────────────────────

    protected override HandleRef BuildWindowCore(HandleRef hwndParent)
    {
        _containerHwnd = CreateWindowEx(
            0, "STATIC", null,
            WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
            0, 0, 1, 1,
            hwndParent.Handle, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);

        if (_excelHwnd != IntPtr.Zero)
            DoAttach();

        return new HandleRef(this, _containerHwnd);
    }

    protected override void DestroyWindowCore(HandleRef hwnd)
    {
        if (_excelHwnd != IntPtr.Zero)
            SetParent(_excelHwnd, IntPtr.Zero);

        if (_containerHwnd != IntPtr.Zero)
        {
            DestroyWindow(_containerHwnd);
            _containerHwnd = IntPtr.Zero;
        }
    }

    // ── 리사이즈: WPF 레이아웃 패스마다 호출 ───────────────────────
    // ArrangeOverride 는 전체화면·최대화·창 크기 변경 모두에서 호출된다.
    // OnRenderSizeChanged 보다 신뢰도가 높고 DPI 스케일도 여기서 처리한다.

    protected override Size ArrangeOverride(Size finalSize)
    {
        var size = base.ArrangeOverride(finalSize);

        // WPF 논리 단위(DIP) → 물리 픽셀 변환 (DPI 125%, 150% 등 대응)
        var (dpiX, dpiY) = GetDpiScale();
        int w = Math.Max(1, (int)(size.Width  * dpiX));
        int h = Math.Max(1, (int)(size.Height * dpiY));

        if (_containerHwnd != IntPtr.Zero)
            MoveWindow(_containerHwnd, 0, 0, w, h, true);
        if (_excelHwnd != IntPtr.Zero)
            MoveWindow(_excelHwnd, 0, 0, w, h, true);

        return size;
    }

    // ── 공개 API ─────────────────────────────────────────────────────

    public void AttachExcel(IntPtr excelHwnd)
    {
        _excelHwnd = excelHwnd;
        if (_containerHwnd != IntPtr.Zero)
            DoAttach();
    }

    // ── 내부 ─────────────────────────────────────────────────────────

    private void DoAttach()
    {
        if (_excelHwnd == IntPtr.Zero || _containerHwnd == IntPtr.Zero) return;

        int style = GetWindowLong(_excelHwnd, GWL_STYLE);
        style &= ~WS_CAPTION;
        style &= ~WS_THICKFRAME;
        style &= ~WS_SYSMENU;
        style &= ~WS_MAXIMIZEBOX;
        style &= ~WS_MINIMIZEBOX;
        style |= WS_CHILD;
        SetWindowLong(_excelHwnd, GWL_STYLE, style);

        SetParent(_excelHwnd, _containerHwnd);

        var (dpiX, dpiY) = GetDpiScale();
        int w = Math.Max(1, (int)(ActualWidth  * dpiX));
        int h = Math.Max(1, (int)(ActualHeight * dpiY));
        MoveWindow(_excelHwnd, 0, 0, w, h, true);
    }

    private (double x, double y) GetDpiScale()
    {
        var source = PresentationSource.FromVisual(this);
        var m = source?.CompositionTarget?.TransformToDevice;
        return (m?.M11 ?? 1.0, m?.M22 ?? 1.0);
    }
}
