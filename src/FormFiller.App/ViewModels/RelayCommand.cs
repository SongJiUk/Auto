using System.Windows.Input;

namespace FormFiller.App.ViewModels;

/// <summary>
/// Unity의 UnityEvent와 비슷한 역할. 버튼 클릭 등 UI 액션을 ViewModel에 연결한다.
/// ICommand 인터페이스: Execute(실행)와 CanExecute(활성화 여부) 두 메서드로 구성.
/// </summary>
public class RelayCommand : ICommand
{
    private readonly Action<object?> _execute;
    private readonly Func<object?, bool>? _canExecute;

    public RelayCommand(Action<object?> execute, Func<object?, bool>? canExecute = null)
    {
        _execute = execute;
        _canExecute = canExecute;
    }

    // CommandManager.RequerySuggested: WPF가 자동으로 CanExecute를 다시 평가하게 한다
    public event EventHandler? CanExecuteChanged
    {
        add    => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }

    public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;
    public void Execute(object? parameter)    => _execute(parameter);
}
