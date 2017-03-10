using System.Windows;

namespace VBASync.WPF {
    internal partial class App {
        private void Application_Startup(object sender, StartupEventArgs e) {
            DispatcherUnhandledException += (s, e2) => {
                MessageBox.Show(e2.Exception.Message, WPF.Properties.Resources.MWTitle,
                    MessageBoxButton.OK, MessageBoxImage.Error);
                e2.Handled = true;
            };
        }
    }
}
