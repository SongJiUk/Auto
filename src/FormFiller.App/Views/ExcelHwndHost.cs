using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace FormFiller.App.Views;

/// <summary>
/// WPF 레이아웃 안의 자리(빈 Win32 창)에 Excel 창을 딱 맞춰 겹쳐 보여 준다.
///
/// [왜 SetParent(자식 창) 방식이 아닌가]
/// Excel 창을 WS_CHILD 로 다른 프로세스 창 안에 넣으면 일부 Excel 버전(2019 빌드 16.0.11328 등)에서
///  - 시트를 바꾸는 순간 통합문서 창을 잃어버려(Windows.Count=0, 리본 회색) 셀 클릭·시트 전환이 안 되고
///  - 붙이는 순서나 숨기기/닫기 순서에 따라 Excel 이 접근 위반으로 종료되며
///  - 창 크기 변경을 놓쳐 예전 크기로 그려지는 일이 생긴다.
/// 그래서 Excel 창은 일반 최상위 창으로 두고, 테두리만 없앤 뒤 이 창을 "소유자"로 지정해
/// 자리 위에 겹쳐 둔다. 소유된 창은 항상 소유자 위에 보이고, 소유자가 최소화되면 같이 숨는다.
///
/// [위치 맞추기] 자리의 크기·위치가 바뀌거나(최대화 포함) 창이 움직일 때마다 Excel 창을 다시 맞춘다.
/// </summary>
internal sealed class ExcelHwndHost : HwndHost
{
    // ── Win32 ────────────────────────────────────────────────────────
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    static extern IntPtr CreateWindowEx(int exStyle, string className, string? title,
        int style, int x, int y, int w, int h,
        IntPtr parent, IntPtr menu, IntPtr instance, IntPtr param);

    [DllImport("user32.dll")] static extern bool DestroyWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool MoveWindow(IntPtr hwnd, int x, int y, int w, int h, bool repaint);
    [DllImport("user32.dll")] static extern int  GetWindowLong(IntPtr hwnd, int index);
    [DllImport("user32.dll")] static extern int  SetWindowLong(IntPtr hwnd, int index, int value);
    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")] static extern IntPtr SetWindowLongPtr64(IntPtr hwnd, int index, IntPtr value);
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int w, int h, uint flags);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll")] static extern bool IsWindow(IntPtr hwnd);

    [StructLayout(LayoutKind.Sequential)]
    struct RECT { public int Left, Top, Right, Bottom; }

    const int GWL_STYLE      = -16;
    const int GWL_EXSTYLE    = -20;
    const int GWLP_HWNDPARENT= -8;    // 최상위 창에서는 "소유자" 를 뜻한다
    const int WS_CHILD       = 0x40000000;
    const int WS_VISIBLE     = 0x10000000;
    const int WS_CLIPCHILDREN= 0x02000000;
    const int WS_CLIPSIBLINGS= 0x04000000;
    const int WS_CAPTION     = 0x00C00000;
    const int WS_THICKFRAME  = 0x00040000;
    const int WS_SYSMENU     = 0x00080000;
    const int WS_MAXIMIZEBOX = 0x00010000;
    const int WS_MINIMIZEBOX = 0x00020000;
    const int WS_EX_APPWINDOW  = 0x00040000;
    const int WS_EX_TOOLWINDOW = 0x00000080;

    static readonly IntPtr HWND_TOP = IntPtr.Zero;
    const uint SWP_NOZORDER     = 0x0004;
    const uint SWP_NOACTIVATE   = 0x0010;
    const uint SWP_FRAMECHANGED = 0x0020;

    private IntPtr  _containerHwnd = IntPtr.Zero;
    private IntPtr  _excelHwnd     = IntPtr.Zero;
    private Window? _owner;

    // ── HwndHost 필수 구현 ──────────────────────────────────────────

    protected override HandleRef BuildWindowCore(HandleRef hwndParent)
    {
        // Excel 이 겹쳐질 "자리" 역할만 하는 빈 창
        _containerHwnd = CreateWindowEx(
            0, "STATIC", null,
            WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
            0, 0, 1, 1,
            hwndParent.Handle, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);

        if (_excelHwnd != IntPtr.Zero)
            Dispatcher.BeginInvoke(DoAttach);

        return new HandleRef(this, _containerHwnd);
    }

    protected override void DestroyWindowCore(HandleRef hwnd)
    {
        Detach();

        if (_containerHwnd != IntPtr.Zero)
        {
            DestroyWindow(_containerHwnd);
            _containerHwnd = IntPtr.Zero;
        }
    }

    // ── 위치·크기 맞추기 ─────────────────────────────────────────────

    protected override Size ArrangeOverride(Size finalSize)
    {
        var size = base.ArrangeOverride(finalSize);
        var (w, h) = ToPixels(size);
        if (_containerHwnd != IntPtr.Zero)
            MoveWindow(_containerHwnd, 0, 0, w, h, true);
        Reposition();
        return size;
    }

    /// <summary>HwndHost 가 자리 창을 옮긴 뒤 호출된다 (레이아웃 변경, 최대화 등).</summary>
    protected override void OnWindowPositionChanged(Rect rcBoundingBox)
    {
        base.OnWindowPositionChanged(rcBoundingBox);
        Reposition();
    }

    private void OnOwnerMoved(object? sender, EventArgs e) => Reposition();

    // ── 공개 API ─────────────────────────────────────────────────────

    public void AttachExcel(IntPtr excelHwnd)
    {
        _excelHwnd = excelHwnd;
        if (_containerHwnd != IntPtr.Zero)
            DoAttach();
    }

    /// <summary>
    /// Excel 창의 소유자 지정을 풀어 독립된 창으로 돌려놓는다.
    /// 통합문서를 닫거나 Excel 을 숨기기 "전에", 그리고 이 창이 닫히기 "전에" 호출해야 한다
    /// (소유자 창이 없어질 때 다른 프로세스의 소유 창이 함께 정리되면서 Excel 이 불안정해질 수 있다).
    /// </summary>
    public void Detach()
    {
        if (_owner != null)
        {
            _owner.LocationChanged -= OnOwnerMoved;
            _owner.StateChanged    -= OnOwnerMoved;
            _owner = null;
        }
        if (_excelHwnd != IntPtr.Zero && IsWindow(_excelHwnd))
            SetOwner(_excelHwnd, IntPtr.Zero);
        _excelHwnd = IntPtr.Zero;
    }

    // ── 내부 ─────────────────────────────────────────────────────────

    private void DoAttach()
    {
        if (_excelHwnd == IntPtr.Zero || _containerHwnd == IntPtr.Zero || _owner != null) return;

        _owner = Window.GetWindow(this);
        if (_owner == null) return;
        var ownerHwnd = new WindowInteropHelper(_owner).Handle;

        // 제목줄·테두리 제거 (창 자체는 최상위 창 그대로 → Excel 이 정상적으로 활성화·입력을 처리한다)
        int style = GetWindowLong(_excelHwnd, GWL_STYLE);
        style &= ~(WS_CAPTION | WS_THICKFRAME | WS_SYSMENU | WS_MAXIMIZEBOX | WS_MINIMIZEBOX);
        SetWindowLong(_excelHwnd, GWL_STYLE, style);

        // 작업 표시줄·Alt+Tab 에 따로 나오지 않게
        int ex = GetWindowLong(_excelHwnd, GWL_EXSTYLE);
        SetWindowLong(_excelHwnd, GWL_EXSTYLE, (ex & ~WS_EX_APPWINDOW) | WS_EX_TOOLWINDOW);

        // 이 창을 소유자로 지정 → 항상 이 창 위에 보이고, 최소화·다른 창 뒤로 가기도 같이 따라간다
        SetOwner(_excelHwnd, ownerHwnd);

        _owner.LocationChanged += OnOwnerMoved;
        _owner.StateChanged    += OnOwnerMoved;

        Reposition(firstTime: true);
    }

    /// <summary>Excel 창을 자리(빈 창)의 화면 위치·크기에 정확히 맞춘다.</summary>
    private void Reposition(bool firstTime = false)
    {
        if (_excelHwnd == IntPtr.Zero || _containerHwnd == IntPtr.Zero || _owner == null) return;
        if (_owner.WindowState == WindowState.Minimized) return;
        if (!GetWindowRect(_containerHwnd, out var r)) return;

        int w = Math.Max(1, r.Right - r.Left), h = Math.Max(1, r.Bottom - r.Top);

        // 처음 붙일 때는 맨 위로 올리고(소유자 위), 스타일 변경을 반영(FRAMECHANGED)한다
        uint flags = SWP_NOACTIVATE | (firstTime ? SWP_FRAMECHANGED : SWP_NOZORDER);
        SetWindowPos(_excelHwnd, HWND_TOP, r.Left, r.Top, w, h, flags);
    }

    private static void SetOwner(IntPtr hwnd, IntPtr owner)
    {
        if (IntPtr.Size == 8) SetWindowLongPtr64(hwnd, GWLP_HWNDPARENT, owner);
        else SetWindowLong(hwnd, GWLP_HWNDPARENT, owner.ToInt32());
    }

    /// <summary>WPF 논리 단위(DIP) → 물리 픽셀 변환 (DPI 125%, 150% 등 대응)</summary>
    private (int w, int h) ToPixels(Size size)
    {
        var source = PresentationSource.FromVisual(this);
        var m = source?.CompositionTarget?.TransformToDevice;
        double dx = m?.M11 ?? 1.0, dy = m?.M22 ?? 1.0;
        return (Math.Max(1, (int)(size.Width * dx)), Math.Max(1, (int)(size.Height * dy)));
    }
}
