using Ookii.Dialogs.Wpf;
using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows;
using VBASync.Localization;

namespace VBASync.WPF
{
    internal sealed class MainViewModel : ViewModelBase, IDisposable
    {
        private readonly Action _onLoadIni;
        private readonly string _lastSessionPath;

        private Model.ActiveSession _activeSession;
        private ChangesViewModel _changes;
        private SessionViewModel _session;
        private SettingsViewModel _settings;

        internal MainViewModel(Model.Startup startup, Action onLoadIni)
        {
            Session = new SessionViewModel
            {
                Action = startup.Action,
                AutoRun = startup.AutoRun,
                FilePath = startup.FilePath,
                FolderPath = startup.FolderPath
            };
            Settings = new SettingsViewModel
            {
                AddNewDocumentsToFile = startup.AddNewDocumentsToFile,
                AfterExtractHookContent = startup.AfterExtractHook?.Content,
                BeforePublishHookContent = startup.BeforePublishHook?.Content,
                DiffTool = startup.DiffTool,
                DiffToolParameters = startup.DiffToolParameters,
                IgnoreEmpty = startup.IgnoreEmpty,
                Language = startup.Language,
                Portable = startup.Portable,
                SearchRepositorySubdirectories = startup.SearchRepositorySubdirectories
            };
            RecentFiles = new BindingList<string>();
            foreach (var recentFile in startup.RecentFiles)
            {
                RecentFiles.Add(recentFile);
            }

            _onLoadIni = onLoadIni;
            _lastSessionPath = startup.LastSessionPath;

            BrowseForSessionCommand = new WpfCommand(v => BrowseForSession());
            LoadLastSessionCommand = new WpfCommand(v => LoadLastSession(), v => LastSessionExists());
            OpenRecentCommand = new WpfCommand(v => LoadRecent((string)v));
            SaveSessionCommand = new WpfCommand(v => SaveSession());

            // promote ListChanged events to PropertyChanged, so the
            // recent files list in the main menu stays up to date
            RecentFiles.ListChanged += (s, e) => OnPropertyChanged(nameof(RecentFiles));
        }

        public Model.ActiveSession ActiveSession => _activeSession;
        public WpfCommand BrowseForSessionCommand { get; }

        public ChangesViewModel Changes
        {
            get => _changes;
            set => SetField(ref _changes, value, nameof(Changes));
        }

        public WpfCommand LoadLastSessionCommand { get; }
        public WpfCommand OpenRecentCommand { get; }
        public BindingList<string> RecentFiles { get; }
        public WpfCommand SaveSessionCommand { get; }

        public SessionViewModel Session
        {
            get => _session;
            set
            {
                SetField(ref _session, value, nameof(Session));
                RefreshActiveSession();
            }
        }

        public SettingsViewModel Settings
        {
            get => _settings;
            set => SetField(ref _settings, value, nameof(Settings));
        }

        public void AddRecentFile(string path)
        {
            var exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var exeDirWithTrailing = exeDir + Path.DirectorySeparatorChar.ToString();
            var displayPath = Path.GetFullPath(path);
            if (displayPath.StartsWith(exeDirWithTrailing))
            {
                displayPath = displayPath.Substring(exeDirWithTrailing.Length);
            }
            if (RecentFiles.Contains(displayPath))
            {
                RecentFiles.Remove(displayPath);
            }
            RecentFiles.Insert(0, displayPath);
            while (RecentFiles.Count > 5)
            {
                RecentFiles.RemoveAt(5);
            }
        }

        public void Dispose()
        {
            _activeSession?.Dispose();
        }

        public void LoadIni(Model.AppIniFile ini)
        {
            Session.Action = ini.GetActionType("General", "ActionType") ?? Model.ActionType.Extract;
            Session.FolderPath = ini.GetString("General", "FolderPath");
            Session.FilePath = ini.GetString("General", "FilePath");
            Settings.AddNewDocumentsToFile = ini.GetBool("General", "AddNewDocumentsToFile") ?? false;
            Settings.IgnoreEmpty = ini.GetBool("General", "IgnoreEmpty") ?? false;
            Settings.AfterExtractHookContent = ini.GetString("Hooks", "AfterExtract");
            Settings.BeforePublishHookContent = ini.GetString("Hooks", "BeforePublish");
            _onLoadIni();
        }

        public void RefreshActiveSession()
        {
            _activeSession?.Dispose();
            _activeSession = new Model.ActiveSession(_session, _settings);
        }

        internal void SaveSession(Stream st, bool saveGlobalSettings, bool saveSessionSettings)
        {
            Func<bool, string> iniBool = b => b ? "true" : "false";
            var sb = new StringBuilder();
            sb.Append("ActionType=").AppendLine(Session.Action == Model.ActionType.Extract ? "Extract" : "Publish");
            sb.Append("FolderPath=\"").Append(Session.FolderPath).AppendLine("\"");
            sb.Append("FilePath=\"").Append(Session.FilePath).AppendLine("\"");
            if (saveSessionSettings)
            {
                sb.Append("AddNewDocumentsToFile=").AppendLine(Settings.AddNewDocumentsToFile ? "true" : "false");
                sb.Append("IgnoreEmpty=").AppendLine(Settings.IgnoreEmpty ? "true" : "false");
                sb.Append("SearchRepositorySubdirectories=").AppendLine(Settings.SearchRepositorySubdirectories ? "true" : "false");
            }
            if (saveGlobalSettings)
            {
                sb.Append("Language=\"").Append(Settings.Language).AppendLine("\"");
                sb.AppendLine("");
                sb.AppendLine("[DiffTool]");
                sb.Append("Path =\"").Append(Settings.DiffTool).AppendLine("\"");
                sb.Append("Parameters=\"").Append(Settings.DiffToolParameters).AppendLine("\"");
            }
            if (saveGlobalSettings
                && (!string.IsNullOrEmpty(Settings.AfterExtractHookContent)
                || !string.IsNullOrEmpty(Settings.BeforePublishHookContent)))
            {
                sb.AppendLine("");
                sb.AppendLine("[Hooks]");
                if (!string.IsNullOrEmpty(Settings.AfterExtractHookContent))
                {
                    sb.Append("AfterExtract=\"").Append(Settings.AfterExtractHookContent).AppendLine("\"");
                }
                if (!string.IsNullOrEmpty(Settings.BeforePublishHookContent))
                {
                    sb.Append("BeforePublish=\"").Append(Settings.BeforePublishHookContent).AppendLine("\"");
                }
            }
            if (saveGlobalSettings && RecentFiles.Count > 0)
            {
                sb.AppendLine("");
                sb.AppendLine("[RecentFiles]");
                var i = 0;
                while (RecentFiles.Count > i)
                {
                    sb.Append((i + 1).ToString(CultureInfo.InvariantCulture))
                        .Append("=\"").Append(RecentFiles[i]).AppendLine("\"");
                    ++i;
                }
            }

            var buf = Encoding.UTF8.GetBytes(sb.ToString());
            st.Write(buf, 0, buf.Length);
        }

        private void BrowseForSession()
        {
            try
            {
                var dlg = new VistaOpenFileDialog
                {
                    Filter = $"{VBASyncResources.MWOpenAllFiles}|*.*|"
                        + $"{VBASyncResources.MWOpenSession}|*.ini",
                    FilterIndex = 2
                };
                if (dlg.ShowDialog() == true)
                {
                    LoadIni(new Model.AppIniFile(dlg.FileName));
                    AddRecentFile(dlg.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, VBASyncResources.MWTitle, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool LastSessionExists() => File.Exists(_lastSessionPath);

        private void LoadLastSession()
        {
            try
            {
                LoadIni(new Model.AppIniFile(_lastSessionPath));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, VBASyncResources.MWTitle, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadRecent(string index)
        {
            try
            {
                var exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var i = int.Parse(index) - 1;
                try
                {
                    var path = RecentFiles[i];
                    if (!Path.IsPathRooted(path))
                    {
                        path = Path.Combine(exeDir, path);
                    }
                    if (!File.Exists(path))
                    {
                        throw new FileNotFoundException();
                    }
                    LoadIni(new Model.AppIniFile(path));
                    AddRecentFile(path);
                }
                catch
                {
                    if (i < RecentFiles.Count)
                    {
                        RecentFiles.RemoveAt(i);
                    }
                    throw;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, VBASyncResources.MWTitle, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveSession()
        {
            try
            {
                var dlg = new VistaSaveFileDialog
                {
                    Filter = $"{VBASyncResources.MWOpenAllFiles}|*.*|"
                        + $"{VBASyncResources.MWOpenSession}|*.ini",
                    FilterIndex = 2
                };
                if (dlg.ShowDialog() == true)
                {
                    var path = dlg.FileName;
                    if (!Path.HasExtension(path))
                    {
                        path += ".ini";
                    }
                    using (var fs = new FileStream(path, FileMode.Create))
                    {
                        SaveSession(fs, false, true);
                    }
                    AddRecentFile(path);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, VBASyncResources.MWTitle, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
