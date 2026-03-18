using System;
using System.Windows.Input;

namespace HotPort.Infrastructure
{
    public class RelayCommand : ICommand
    {
        private readonly Action<object?> execute;
        private readonly Predicate<object?>? canExecute;

        public RelayCommand(Action execute)
            : this(_ => execute(), null)
        {
        }

        public RelayCommand(Action execute, Func<bool> canExecute)
            : this(_ => execute(), _ => canExecute())
        {
        }

        public RelayCommand(Action<object?> execute)
            : this(execute, null)
        {
        }

        public RelayCommand(Action<object?> execute, Predicate<object?>? canExecute)
        {
            this.execute = execute ?? throw new ArgumentNullException(nameof(execute));
            this.canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        public bool CanExecute(object? parameter)
        {
            return canExecute?.Invoke(parameter) ?? true;
        }

        public void Execute(object? parameter)
        {
            execute(parameter);
        }

        public void RaiseCanExecuteChanged()
        {
            CommandManager.InvalidateRequerySuggested();
        }
    }
}