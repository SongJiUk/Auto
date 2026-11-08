using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace FormFiller.App.Views;

/// <summary>
/// WPF HwndHost 를 상속해 Excel 창을 WPF 레이아웃 안에 직접 임베드한다.
/// WindowsFormsHost 를 쓰지 않으므로 Office COM 메시지 루프와 충돌이 없다.
///
/// [원리] BuildWindowCore 에서 중간 컨테이너 Win32 창을 만들고,
///        AttachExcel 호출 시 Excel 창을 그 컨테이너의 자식으로 reparent 한다.
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

    // ── 상태 ─────────────────────────────────────────────────────────
    private IntPtr _containerHwnd = IntPtr.Zero;
    private IntPtr _excelHwnd     = IntPtr.Zero;

    // ── HwndHost 필수 구현 ──────────────────────────────────────────

    protected override HandleRef BuildWindowCore(HandleRef hwndParent)
    {
        // hwndParent 의 자식 컨테이너 창 생성
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
        // Excel 창은 여기서 닫지 않음 — CleanupExcel 에서 처리
        if (_excelHwnd != IntPtr.Zero)
            SetParent(_excelHwnd, IntPtr.Zero);

        if (_containerHwnd != IntPtr.Zero)
        {
            DestroyWindow(_containerHwnd);
            _containerHwnd = IntPtr.Zero;
        }
    }

    // ── 공개 API ─────────────────────────────────────────────────────

    /// <summary>Excel 창 핸들을 받아 컨테이너 안으로 reparent 한다.</summary>
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

        // 타이틀바·테두리 제거 후 자식 창으로 변환
        int style = GetWindowLong(_excelHwnd, GWL_STYLE);
        style &= ~WS_CAPTION;
        style &= ~WS_THICKFRAME;
        style &= ~WS_SYSMENU;
        style &= ~WS_MAXIMIZEBOX;
        style &= ~WS_MINIMIZEBOX;
        style |= WS_CHILD;
        SetWindowLong(_excelHwnd, GWL_STYLE, style);

        SetParent(_excelHwnd, _containerHwnd);

        int w = Math.Max(1, (int)ActualWidth);
        int h = Math.Max(1, (int)ActualHeight);
        MoveWindow(_excelHwnd, 0, 0, w, h, true);
    }

    protected override void OnRenderSizeChanged(SizeChangedInfo info)
    {
        base.OnRenderSizeChanged(info);
        int w = Math.Max(1, (int)info.NewSize.Width);
        int h = Math.Max(1, (int)info.NewSize.Height);

        if (_containerHwnd != IntPtr.Zero)
            MoveWindow(_containerHwnd, 0, 0, w, h, true);
        if (_excelHwnd != IntPtr.Zero)
            MoveWindow(_excelHwnd, 0, 0, w, h, true);
    }
}
