using System;
using System.Windows.Input;

namespace VBASync.WPF
{
    internal class WpfCommand : ICommand
    {
        private readonly Action<dynamic> _command;

        internal WpfCommand(Action<dynamic> command) => _command = command;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(dynamic parameter) => true;
        public void Execute(dynamic parameter) => _command(parameter);
    }
}
