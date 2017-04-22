using System;
using System.Windows.Input;

namespace VBASync.WPF
{
    internal class WpfCommand : ICommand
    {
        private readonly Func<dynamic, bool> _canExecute;
        private readonly Action<dynamic> _command;

        internal WpfCommand(Action<dynamic> command, Func<dynamic, bool> canExecute = null)
        {
            _command = command;
            _canExecute = canExecute ?? (v => true);
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(dynamic parameter) => _canExecute(parameter);
        public void Execute(dynamic parameter) => _command(parameter);
        public void OnExecuteChanged() => CanExecuteChanged?.Invoke(this, new EventArgs());
    }
}
