using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using FormFiller.Core.Config;

namespace FormFiller.App;

public partial class App : System.Windows.Application
{
    /// <summary>
    /// 예상하지 못한 오류가 나도 프로그램이 그냥 꺼지지 않게 하고, logs 폴더에 기록을 남긴다.
    /// </summary>
    protected override void OnStartup(StartupEventArgs e)
    {
        AppLog.Info($"프로그램 시작 (PC: {System.Environment.MachineName}, 폴더: {System.AppDomain.CurrentDomain.BaseDirectory})");

        DispatcherUnhandledException += OnUiException;
        System.AppDomain.CurrentDomain.UnhandledException += (_, args) =>
            AppLog.Error("처리하지 못한 오류 (프로그램 종료)", args.ExceptionObject as System.Exception);
        TaskScheduler.UnobservedTaskException += (_, args) =>
        {
            AppLog.Error("뒤에서 돌던 작업의 오류", args.Exception);
            args.SetObserved();
        };

        base.OnStartup(e);
    }

    protected override void OnExit(ExitEventArgs e)
    {
        AppLog.Info("프로그램 종료");
        base.OnExit(e);
    }

    private void OnUiException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        AppLog.Error("화면 처리 중 오류", e.Exception);
        MessageBox.Show(
            $"문제가 생겼습니다. 하던 작업을 확인해 주세요.\n\n{e.Exception.Message}\n\n" +
            $"자세한 내용은 기록 파일에 남겼습니다:\n{AppLog.LogPath}",
            "오류", MessageBoxButton.OK, MessageBoxImage.Error);
        e.Handled = true;   // 프로그램을 끄지 않고 계속 쓴다
    }
}
