using NUnit.Framework;
using System;
using VBASync.Model;
using VBASync.Tests.Mocks;

namespace VBASync.Tests.UnitTests
{
    [TestFixture]
    public class StartupTests
    {
        [Test]
        public void StartupAllArgsExtract()
        {
            var startup = new Startup(MakeWindowsHook, s => ThrowUnexpectedIniRequest());
            startup.ProcessArgs(new[] { "-x", "-f", @"C:\Path\To\File", "-d", @"C:\Path\To\Folder",
                "-r", "-a", "-i", "-h", @"MyAwesomeHook.bat ""{TargetDir}""" });

            Assert.That(startup.Action, Is.EqualTo(ActionType.Extract));
            Assert.That(startup.AddNewDocumentsToFile, Is.EqualTo(true));
            Assert.That(startup.AfterExtractHook.Content, Is.EqualTo(@"MyAwesomeHook.bat ""{TargetDir}"""));
            Assert.That(startup.AutoRun, Is.EqualTo(true));
            Assert.That(startup.FilePath, Is.EqualTo(@"C:\Path\To\File"));
            Assert.That(startup.FolderPath, Is.EqualTo(@"C:\Path\To\Folder"));
            Assert.That(startup.IgnoreEmpty, Is.EqualTo(true));
        }

        [Test]
        public void StartupAllArgsPublish()
        {
            var startup = new Startup(MakeWindowsHook, s => ThrowUnexpectedIniRequest());
            startup.ProcessArgs(new[] { "-p", "-f", @"C:\Path\To\File", "-d", @"C:\Path\To\Folder",
                "-r", "-a", "-i", "-h", @"MyAwesomeHook.bat ""{TargetDir}""" });

            Assert.That(startup.Action, Is.EqualTo(ActionType.Publish));
            Assert.That(startup.AddNewDocumentsToFile, Is.EqualTo(true));
            Assert.That(startup.BeforePublishHook.Content, Is.EqualTo(@"MyAwesomeHook.bat ""{TargetDir}"""));
            Assert.That(startup.AutoRun, Is.EqualTo(true));
            Assert.That(startup.FilePath, Is.EqualTo(@"C:\Path\To\File"));
            Assert.That(startup.FolderPath, Is.EqualTo(@"C:\Path\To\Folder"));
            Assert.That(startup.IgnoreEmpty, Is.EqualTo(true));
        }

        [Test]
        public void StartupAutoExtract()
        {
            var startup = new Startup(MakeWindowsHook, s => ThrowUnexpectedIniRequest());

            var generalIni = new AppIniFile();
            generalIni.PumpValue("General", "Language", "en");
            generalIni.PumpValue("General", "ActionType", "Extract");
            generalIni.PumpValue("General", "FolderPath", @"C:\Path\To\LastFolder");
            generalIni.PumpValue("General", "FilePath", @"C:\Path\To\LastFile");
            generalIni.PumpValue("DiffTool", "Path", @"C:\Path\To\DiffTool");
            startup.ProcessIni(generalIni, false);

            var sessionIni = new AppIniFile();
            sessionIni.PumpValue("General", "ActionType", "Publish");
            sessionIni.PumpValue("General", "FolderPath", @"C:\Path\To\RealFolder");
            sessionIni.PumpValue("General", "FilePath", @"C:\Path\To\RealFile");
            startup.ProcessIni(sessionIni, true);

            startup.ProcessArgs(new[] { "-r", "-x" });

            Assert.That(startup.Language, Is.EqualTo("en"));
            Assert.That(startup.Action, Is.EqualTo(ActionType.Extract));
            Assert.That(startup.FolderPath, Is.EqualTo(@"C:\Path\To\RealFolder"));
            Assert.That(startup.FilePath, Is.EqualTo(@"C:\Path\To\RealFile"));
            Assert.That(startup.DiffTool, Is.EqualTo(@"C:\Path\To\DiffTool"));
            Assert.That(startup.RecentFiles.Count, Is.EqualTo(0));
            Assert.That(startup.AutoRun, Is.EqualTo(true));
        }

        [Test]
        public void StartupAutoPublishDueIniOverride()
        {
            var startup = new Startup(MakeWindowsHook, AllowOnlyPublisherIniRequest);

            var generalIni = new AppIniFile();
            generalIni.PumpValue("General", "Language", "en");
            generalIni.PumpValue("General", "ActionType", "Extract");
            generalIni.PumpValue("General", "FolderPath", @"C:\Path\To\LastFolder");
            generalIni.PumpValue("General", "FilePath", @"C:\Path\To\LastFile");
            generalIni.PumpValue("DiffTool", "Path", @"C:\Path\To\DiffTool");
            startup.ProcessIni(generalIni, false);

            var sessionIni = new AppIniFile();
            sessionIni.PumpValue("General", "ActionType", "Publish");
            sessionIni.PumpValue("General", "FolderPath", @"C:\Path\To\RealFolder");
            sessionIni.PumpValue("General", "FilePath", @"C:\Path\To\RealFile");
            startup.ProcessIni(sessionIni, true);

            startup.ProcessArgs(new[] { "-r", "-x", "Publisher.ini" });

            Assert.That(startup.Language, Is.EqualTo("en"));
            Assert.That(startup.Action, Is.EqualTo(ActionType.Publish));
            Assert.That(startup.FolderPath, Is.EqualTo(@"C:\Path\To\RealFolder"));
            Assert.That(startup.FilePath, Is.EqualTo(@"C:\Path\To\RealFile"));
            Assert.That(startup.DiffTool, Is.EqualTo(@"C:\Path\To\DiffTool"));
            Assert.That(startup.RecentFiles.Count, Is.EqualTo(0));
            Assert.That(startup.AutoRun, Is.EqualTo(true));
        }

        [Test]
        public void StartupDefaultDiffToolParameters()
        {
            var startup = new Startup();
            Assert.That(startup.DiffToolParameters, Is.EqualTo(@"""{OldFile}"" ""{NewFile}"""));
        }

        [Test]
        public void StartupInteractiveNoSessionIni()
        {
            var startup = new Startup(MakeWindowsHook, s => ThrowUnexpectedIniRequest());

            var ini = new AppIniFile();
            ini.PumpValue("General", "Language", "en");
            ini.PumpValue("General", "ActionType", "Extract");
            ini.PumpValue("General", "FolderPath", @"C:\Path\To\LastFolder");
            ini.PumpValue("General", "FilePath", @"C:\Path\To\LastFile");
            ini.PumpValue("DiffTool", "Path", @"C:\Path\To\DiffTool");
            ini.PumpValue("RecentFiles", "1", @"C:\Path\To\RecentFile1");
            ini.PumpValue("RecentFiles", "2", @"C:\Path\To\RecentFile2");
            startup.ProcessIni(ini, false);

            startup.ProcessIni(new AppIniFile(), true);

            startup.ProcessArgs(new string[0]);

            Assert.That(startup.Language, Is.EqualTo("en"));
            Assert.That(startup.Action, Is.EqualTo(ActionType.Extract));
            Assert.That(startup.FolderPath, Is.Null);
            Assert.That(startup.FilePath, Is.Null);
            Assert.That(startup.DiffTool, Is.EqualTo(@"C:\Path\To\DiffTool"));
            Assert.That(startup.RecentFiles.Count, Is.EqualTo(2));
            Assert.That(startup.RecentFiles[0], Is.EqualTo(@"C:\Path\To\RecentFile1"));
            Assert.That(startup.RecentFiles[1], Is.EqualTo(@"C:\Path\To\RecentFile2"));
            Assert.That(startup.AutoRun, Is.EqualTo(false));
        }

        [Test]
        public void StartupInteractiveWithSessionIni()
        {
            var startup = new Startup(MakeWindowsHook, s => ThrowUnexpectedIniRequest());

            var generalIni = new AppIniFile();
            generalIni.PumpValue("General", "Language", "en");
            generalIni.PumpValue("General", "ActionType", "Extract");
            generalIni.PumpValue("General", "FolderPath", @"C:\Path\To\LastFolder");
            generalIni.PumpValue("General", "FilePath", @"C:\Path\To\LastFile");
            generalIni.PumpValue("DiffTool", "Path", @"C:\Path\To\DiffTool");
            startup.ProcessIni(generalIni, false);

            var sessionIni = new AppIniFile();
            sessionIni.PumpValue("General", "ActionType", "Publish");
            sessionIni.PumpValue("General", "FolderPath", @"C:\Path\To\RealFolder");
            sessionIni.PumpValue("General", "FilePath", @"C:\Path\To\RealFile");
            startup.ProcessIni(sessionIni, true);

            startup.ProcessArgs(new string[0]);

            Assert.That(startup.Language, Is.EqualTo("en"));
            Assert.That(startup.Action, Is.EqualTo(ActionType.Publish));
            Assert.That(startup.FolderPath, Is.EqualTo(@"C:\Path\To\RealFolder"));
            Assert.That(startup.FilePath, Is.EqualTo(@"C:\Path\To\RealFile"));
            Assert.That(startup.DiffTool, Is.EqualTo(@"C:\Path\To\DiffTool"));
            Assert.That(startup.RecentFiles.Count, Is.EqualTo(0));
            Assert.That(startup.AutoRun, Is.EqualTo(false));
        }

        private AppIniFile AllowOnlyPublisherIniRequest(string path)
        {
            if (string.Equals(path, "Publisher.ini", StringComparison.InvariantCultureIgnoreCase))
            {
                var ini = new AppIniFile();
                ini.PumpValue("General", "ActionType", "Publish");
                return ini;
            }
            return ThrowUnexpectedIniRequest();
        }

        private Hook MakeWindowsHook(string content) => new Hook(new WindowsFakeSystemOperations(), content);

        private AppIniFile ThrowUnexpectedIniRequest()
        {
            throw new ApplicationException("Encountered an unexpected request for a .ini file!");
        }
    }
}
