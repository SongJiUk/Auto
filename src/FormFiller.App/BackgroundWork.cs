using System.Threading;
using System.Windows.Threading;

namespace FormFiller.App;

/// <summary>
/// 오래 걸리는 일(파일 읽기·쓰기, 사진 줄이기)을 화면 스레드 밖에서 돌려 창이 멈추지 않게 한다.
/// WPF 사진 처리(BitmapFrame, RenderTargetBitmap)는 STA 스레드가 필요해서 Task.Run 대신 STA 스레드를 만든다.
/// 결과로 받은 사진(BitmapSource)은 Freeze 해야 화면 스레드에서 쓸 수 있다.
/// </summary>
public static class BackgroundWork
{
    // 드라이브 폴더에서 사진을 한꺼번에 많이 읽으면 오히려 느려져서 동시에 2개까지만
    private static readonly SemaphoreSlim Slots = new(2);

    public static async Task<T> RunSta<T>(Func<T> work)
    {
        await Slots.WaitAsync();
        try
        {
            var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            var thread = new Thread(() =>
            {
                try { tcs.SetResult(work()); }
                catch (Exception ex) { tcs.SetException(ex); }
                finally { Dispatcher.FromThread(Thread.CurrentThread)?.InvokeShutdown(); }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
            return await tcs.Task;
        }
        finally { Slots.Release(); }
    }
}
