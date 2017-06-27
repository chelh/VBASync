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

                var generalIni = new Model.AppIniFile(Path.Combine(exeDir, "VBASync.ini"));
                if (!string.Equals(exeBaseName, "VBASync", StringComparison.InvariantCultureIgnoreCase))
                {
                    generalIni.AddFile(Path.Combine(exeDir, exeBaseName + ".ini"));
                }

                var appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                var lastSessionPath = generalIni.GetBool("General", "Portable") ?? false
                    ? Path.Combine(exeDir, "LastSession.ini")
                    : Path.Combine(appDataDir, "VBA Sync Tool", "LastSession.ini");
                generalIni.AddFile(lastSessionPath);

                var startup = new Model.Startup { LastSessionPath = lastSessionPath };
                startup.ProcessIni(generalIni, false); // don't allow loading session settings from these .ini files

                var sessionIni = new Model.AppIniFile(Path.Combine(Environment.CurrentDirectory, "VBASync.ini"));
                sessionIni.AddFile(Path.Combine(Environment.CurrentDirectory, "VBASync.ini.local"));
                if (!string.Equals(exeBaseName, "VBASync", StringComparison.InvariantCultureIgnoreCase))
                {
                    sessionIni.AddFile(Path.Combine(Environment.CurrentDirectory, exeBaseName + ".ini"));
                    sessionIni.AddFile(Path.Combine(Environment.CurrentDirectory, exeBaseName + ".ini.local"));
                }
                startup.ProcessIni(sessionIni, true);

                startup.ProcessArgs(args);

                if (!string.IsNullOrEmpty(startup.Language))
                {
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo(startup.Language);
                }

                if (startup.AutoRun)
                {
                    using (var actor = new Model.ActiveSession(startup, startup))
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
