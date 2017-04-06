using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
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

            try
            {
                var exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var exeBaseName = Process.GetCurrentProcess().ProcessName;
                var ini = new Model.AppIniFile(Path.Combine(exeDir, "VBASync.ini"));
                ini.AddFile(Path.Combine(exeDir, exeBaseName + ".ini"));
                var appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                var lastSessionPath = ini.GetBool("General", "Portable") ?? false
                    ? Path.Combine(exeDir, "LastSession.ini")
                    : Path.Combine(appDataDir, "VBA Sync Tool", "LastSession.ini");
                ini.AddFile(lastSessionPath);

                // don't persist these settings
                ini.Delete("General", "ActionType");
                ini.Delete("General", "FolderPath");
                ini.Delete("General", "FilePath");
                ini.Delete("General", "AutoRun");

                if (exeDir != Environment.CurrentDirectory)
                {
                    ini.AddFile(Path.Combine(Environment.CurrentDirectory, "VBASync.ini"));
                    ini.AddFile(Path.Combine(Environment.CurrentDirectory, exeBaseName + ".ini"));
                }
                var args = Environment.GetCommandLineArgs();
                var autoRunSwitch = false;
                Model.ActionType? actionSwitch = null;
                string filePathSwitch = null;
                string folderPathSwitch = null;
                for (var i = 1; i < args.Length; ++i)
                {
                    switch (args[i].ToUpperInvariant())
                    {
                        case "-R":
                        case "/R":
                            autoRunSwitch = true;
                            break;
                        case "-X":
                        case "/X":
                            actionSwitch = Model.ActionType.Extract;
                            break;
                        case "-P":
                        case "/P":
                            actionSwitch = Model.ActionType.Publish;
                            break;
                        case "-F":
                        case "/F":
                            filePathSwitch = args[++i];
                            break;
                        case "-D":
                        case "/D":
                            folderPathSwitch = args[++i];
                            break;
                        default:
                            ini.AddFile(args[i]);
                            break;
                    }
                }

                var initialSession = new WPF.MainViewModel
                {
                    Session = new WPF.SessionViewModel
                    {
                        Action = actionSwitch ?? ini.GetActionType("General", "ActionType") ?? Model.ActionType.Extract,
                        AutoRun = autoRunSwitch || (ini.GetBool("General", "AutoRun") ?? false),
                        FilePath = filePathSwitch ?? ini.GetString("General", "FilePath"),
                        FolderPath = folderPathSwitch ?? ini.GetString("General", "FolderPath")
                    },
                    Settings = new WPF.SettingsViewModel
                    {
                        DiffTool = ini.GetString("DiffTool", "Path"),
                        DiffToolParameters = ini.GetString("DiffTool", "Parameters") ?? "\"{OldFile}\" \"{NewFile}\"",
                        Language = ini.GetString("General", "Language"),
                        Portable = ini.GetBool("General", "Portable") ?? false
                    }
                };
                var j = 1;
                while (j <= 5 && !string.IsNullOrEmpty(ini.GetString("RecentFiles", j.ToString(CultureInfo.InvariantCulture))))
                {
                    initialSession.RecentFiles.Add(ini.GetString("RecentFiles", j.ToString(CultureInfo.InvariantCulture)));
                    ++j;
                }
                if (!string.IsNullOrEmpty(initialSession.Settings.Language))
                {
                    VBASyncResources.Culture = new CultureInfo(initialSession.Settings.Language);
                }
                var mw = new WPF.MainWindow(initialSession);
                mw.Show();
                if (!File.Exists(lastSessionPath) && !initialSession.Session.AutoRun)
                {
                    mw.SettingsMenu_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, VBASyncResources.MWTitle, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
