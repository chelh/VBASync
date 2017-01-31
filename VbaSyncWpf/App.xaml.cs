using System.Windows;

namespace VbaSync {
    internal partial class App {
        private void Application_Startup(object sender, StartupEventArgs e) {
            DispatcherUnhandledException += (s, e2) => {
                MessageBox.Show(e2.Exception.Message, VbaSync.Properties.Resources.MWTitle, MessageBoxButton.OK, MessageBoxImage.Error);
                e2.Handled = true;
            };
        }
    }
}
