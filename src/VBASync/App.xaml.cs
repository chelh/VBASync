using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows;
using VBASync.Localization;

namespace VBASync
{
    internal partial class App
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            DispatcherUnhandledException += (s, e2) => {
                MessageBox.Show(e2.Exception.Message, VBASyncResources.MWTitle,
                    MessageBoxButton.OK, MessageBoxImage.Error);
                e2.Handled = true;
            };

            var ini = new Model.AppIniFile(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                "VBASync.ini"));
            var lastSessionPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "VBA Sync Tool", "LastSession.ini");
            ini.AddFile(lastSessionPath);

            // don't persist these settings
            ini.Delete("General", "ActionType");
            ini.Delete("General", "FolderPath");
            ini.Delete("General", "FilePath");
            ini.Delete("General", "AutoRun");

            ini.AddFile(Path.Combine(Environment.CurrentDirectory, "VBASync.ini"));
            ini.AddFile(Path.Combine(Environment.CurrentDirectory, Process.GetCurrentProcess().ProcessName + ".ini"));
            var args = Environment.GetCommandLineArgs();
            var autoRunSwitch = false;
            for (var i = 1; i < args.Length; i++)
            {
                switch (args[i].ToUpperInvariant())
                {
                    case "-R":
                    case "/R":
                        autoRunSwitch = true;
                        break;
                    default:
                        ini.AddFile(args[i]);
                        break;
                }
            }

            var initialSession = new WPF.MainViewModel
            {
                Action = ini.GetActionType("General", "ActionType") ?? Model.ActionType.Extract,
                AutoRun = autoRunSwitch || (ini.GetBool("General", "AutoRun") ?? false),
                DiffTool = ini.GetString("DiffTool", "Path"),
                DiffToolParameters = ini.GetString("DiffTool", "Parameters") ?? "\"{OldFile}\" \"{NewFile}\"",
                FilePath = ini.GetString("General", "FilePath"),
                FolderPath = ini.GetString("General", "FolderPath"),
                Language = ini.GetString("General", "Language")
            };
            if (!string.IsNullOrEmpty(initialSession.Language))
            {
                VBASyncResources.Culture = new CultureInfo(initialSession.Language);
            }
            var mw = new WPF.MainWindow(initialSession);
            mw.Show();
            if (!File.Exists(lastSessionPath))
            {
                mw.SettingsMenu_Click(null, null);
            }
        }
    }
}
