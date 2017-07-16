using System;
using System.Collections.Generic;
using System.Globalization;

namespace VBASync.Model
{
    public class Startup : ISession, ISessionSettings
    {
        private readonly Func<string, AppIniFile> _appIniFileFactory;
        private readonly Func<string, Hook> _hookFactory;

        public Startup() : this(s => new Hook(s), s => new AppIniFile(s))
        {
        }

        internal Startup(Func<string, Hook> hookFactory, Func<string, AppIniFile> appIniFileFactory)
        {
            _hookFactory = hookFactory;
            _appIniFileFactory = appIniFileFactory;
        }

        public ActionType Action { get; set; }
        public bool AddNewDocumentsToFile { get; set; }
        public Hook AfterExtractHook { get; set; }
        public bool AutoRun { get; set; }
        public Hook BeforePublishHook { get; set; }
        public bool DeleteDocumentsFromFile { get; set; }
        public string DiffTool { get; set; }
        public string DiffToolParameters { get; set; } = "\"{OldFile}\" \"{NewFile}\"";
        public string FilePath { get; set; }
        public string FolderPath { get; set; }
        public bool IgnoreEmpty { get; set; }
        public string Language { get; set; }
        public string LastSessionPath { get; set; }
        public bool Portable { get; set; }
        public List<string> RecentFiles { get; } = new List<string>();
        public bool SearchRepositorySubdirectories { get; set; }

        public void ProcessArgs(string[] args)
        {
            for (var i = 0; i < args.Length; ++i)
            {
                switch (args[i].ToUpperInvariant())
                {
                    case "-R":
                    case "/R":
                        AutoRun = true;
                        break;
                    case "-X":
                    case "/X":
                        Action = ActionType.Extract;
                        break;
                    case "-P":
                    case "/P":
                        Action = ActionType.Publish;
                        break;
                    case "-F":
                    case "/F":
                        FilePath = args[++i];
                        break;
                    case "-D":
                    case "/D":
                        FolderPath = args[++i];
                        break;
                    case "-A":
                    case "/A":
                        AddNewDocumentsToFile = true;
                        break;
                    case "-I":
                    case "/I":
                        IgnoreEmpty = true;
                        break;
                    case "-H":
                    case "/H":
                        if (Action == ActionType.Publish)
                        {
                            BeforePublishHook = _hookFactory(args[++i]);
                        }
                        else
                        {
                            AfterExtractHook = _hookFactory(args[++i]);
                        }
                        break;
                    case "-E":
                    case "/E":
                        DeleteDocumentsFromFile = true;
                        break;
                    case "-U":
                    case "/U":
                        SearchRepositorySubdirectories = true;
                        break;
                    default:
                        ProcessIni(_appIniFileFactory(args[i]), true);
                        break;
                }
            }
        }

        public void ProcessIni(AppIniFile ini, bool allowSessionSettings)
        {
            var iniAction = ini.GetActionType("General", "ActionType");
            var iniAddNewDocumentsToFile = ini.GetBool("General", "AddNewDocumentsToFile");
            var iniAfterExtractHook = ini.GetString("Hooks", "AfterExtract");
            var iniAutoRun = ini.GetBool("General", "AutoRun");
            var iniBeforePublishHook = ini.GetString("Hooks", "BeforePublish");
            var iniDeleteDocumentsFromFile = ini.GetBool("General", "DeleteDocumentsFromFile");
            var iniDiffTool = ini.GetString("DiffTool", "Path");
            var iniDiffToolParameters = ini.GetString("DiffTool", "Parameters");
            var iniIgnoreEmpty = ini.GetBool("General", "IgnoreEmpty");
            var iniFilePath = ini.GetString("General", "FilePath");
            var iniFolderPath = ini.GetString("General", "FolderPath");
            var iniLanguage = ini.GetString("General", "Language");
            var iniPortable = ini.GetBool("General", "Portable");
            var iniSearchSubdirectories = ini.GetBool("General", "SearchRepositorySubdirectories");

            if (iniAction.HasValue && allowSessionSettings)
            {
                Action = iniAction.Value;
            }

            if (iniAddNewDocumentsToFile.HasValue && allowSessionSettings)
            {
                AddNewDocumentsToFile = iniAddNewDocumentsToFile.Value;
            }

            if (iniAfterExtractHook != null && allowSessionSettings)
            {
                AfterExtractHook = _hookFactory(iniAfterExtractHook);
            }

            if (iniAutoRun.HasValue && allowSessionSettings)
            {
                AutoRun = iniAutoRun.Value;
            }

            if (iniBeforePublishHook != null && allowSessionSettings)
            {
                BeforePublishHook = _hookFactory(iniBeforePublishHook);
            }

            if (iniDeleteDocumentsFromFile.HasValue && allowSessionSettings)
            {
                DeleteDocumentsFromFile = iniDeleteDocumentsFromFile.Value;
            }

            if (iniDiffTool != null)
            {
                DiffTool = iniDiffTool;
            }

            if (iniDiffToolParameters != null)
            {
                DiffToolParameters = iniDiffToolParameters;
            }

            if (iniIgnoreEmpty.HasValue && allowSessionSettings)
            {
                IgnoreEmpty = iniIgnoreEmpty.Value;
            }

            if (iniFilePath != null && allowSessionSettings)
            {
                FilePath = iniFilePath;
            }

            if (iniFolderPath != null && allowSessionSettings)
            {
                FolderPath = iniFolderPath;
            }

            if (iniLanguage != null)
            {
                Language = iniLanguage;
            }

            if (iniPortable.HasValue)
            {
                Portable = iniPortable.Value;
            }

            if (iniSearchSubdirectories.HasValue)
            {
                SearchRepositorySubdirectories = iniSearchSubdirectories.Value;
            }

            if (ini.GetString("RecentFiles", "1") != null)
            {
                RecentFiles.Clear();
                var j = 1;
                while (j <= 5 && !string.IsNullOrEmpty(ini.GetString("RecentFiles", j.ToString(CultureInfo.InvariantCulture))))
                {
                    RecentFiles.Add(ini.GetString("RecentFiles", j.ToString(CultureInfo.InvariantCulture)));
                    ++j;
                }
            }
        }
    }
}
