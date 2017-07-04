using System.Windows;
using VBASync.Localization;

namespace VBASync.WPF
{
    public static class WpfManager
    {
        public static void RunWpf(Model.Startup startup, bool showSettingsWindow)
        {
            var app = new Application();
            app.DispatcherUnhandledException += (s, e) => {
                MessageBox.Show(e.Exception.Message, VBASyncResources.MWTitle,
                    MessageBoxButton.OK, MessageBoxImage.Error);
                e.Handled = true;
            };
            using (var mw = new MainWindow(startup))
            {
                mw.Show();
                if (showSettingsWindow)
                {
                    mw.SettingsMenu_Click(null, null);
                }
                app.Run();
            }
        }
    }
}
