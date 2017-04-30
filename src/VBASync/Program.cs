using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using VBASync.Localization;
using Forms = System.Windows.Forms;

namespace VBASync
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            try
            {
                var exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var exeBaseName = Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().Location);
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

                var startup = new Model.Startup
                {
                    Action = actionSwitch ?? ini.GetActionType("General", "ActionType") ?? Model.ActionType.Extract,
                    AutoRun = autoRunSwitch || (ini.GetBool("General", "AutoRun") ?? false),
                    FilePath = filePathSwitch ?? ini.GetString("General", "FilePath"),
                    FolderPath = folderPathSwitch ?? ini.GetString("General", "FolderPath"),
                    DiffTool = ini.GetString("DiffTool", "Path"),
                    DiffToolParameters = ini.GetString("DiffTool", "Parameters") ?? "\"{OldFile}\" \"{NewFile}\"",
                    Language = ini.GetString("General", "Language"),
                    Portable = ini.GetBool("General", "Portable") ?? false
                };
                var j = 1;
                while (j <= 5 && !string.IsNullOrEmpty(ini.GetString("RecentFiles", j.ToString(CultureInfo.InvariantCulture))))
                {
                    startup.RecentFiles.Add(ini.GetString("RecentFiles", j.ToString(CultureInfo.InvariantCulture)));
                    ++j;
                }
                if (!string.IsNullOrEmpty(startup.Language))
                {
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo(startup.Language);
                }
                if (startup.AutoRun)
                {
                    using (var actor = new Model.ActiveSession(startup))
                    {
                        actor.Apply(actor.GetPatches().ToList());
                    }
                }
                else
                {
                    try
                    {
                        Assembly.Load("VBASync.WPF")
                            .GetType("VBASync.WPF.WpfManager")
                            .GetMethod("RunWpf", BindingFlags.Public | BindingFlags.Static)
                            .Invoke(null, new object[] { startup, !File.Exists(lastSessionPath) });
                    }
                    catch
                    {
                        throw new ApplicationException(VBASyncResources.ErrorCannotLoadGUI);
                    }
                }
            }
            catch (Exception ex)
            {
                Forms.MessageBox.Show(ex.Message, VBASyncResources.MWTitle,
                    Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Error);
            }
        }
    }
}
