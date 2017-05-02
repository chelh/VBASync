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
    internal class MainViewModel : ViewModelBase
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
                DiffTool = startup.DiffTool,
                DiffToolParameters = startup.DiffToolParameters,
                IgnoreEmpty = startup.IgnoreEmpty,
                Language = startup.Language,
                Portable = startup.Portable
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
            var exeDirWithTrailing = exeDir + Path.DirectorySeparatorChar;
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

        public void LoadIni(Model.AppIniFile ini)
        {
            Session.Action = ini.GetActionType("General", "ActionType") ?? Session.Action;
            Session.FolderPath = ini.GetString("General", "FolderPath") ?? Session.FolderPath;
            Session.FilePath = ini.GetString("General", "FilePath") ?? Session.FilePath;
            Settings.AddNewDocumentsToFile = ini.GetBool("General", "AddNewDocumentsToFile")
                ?? Settings.AddNewDocumentsToFile;
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
            sb.AppendLine("ActionType=" + (Session.Action == Model.ActionType.Extract ? "Extract" : "Publish"));
            sb.AppendLine($"FolderPath=\"{Session.FolderPath}\"");
            sb.AppendLine($"FilePath=\"{Session.FilePath}\"");
            if (saveSessionSettings)
            {
                sb.AppendLine($"AddNewDocumentsToFile={iniBool(Settings.AddNewDocumentsToFile)}");
                sb.AppendLine($"IgnoreEmpty={iniBool(Settings.IgnoreEmpty)}");
            }
            if (saveGlobalSettings)
            {
                sb.AppendLine($"Language=\"{Settings.Language}\"");
                sb.AppendLine("");
                sb.AppendLine("[DiffTool]");
                sb.AppendLine($"Path =\"{Settings.DiffTool}\"");
                sb.AppendLine($"Parameters=\"{Settings.DiffToolParameters}\"");
                if (RecentFiles.Count > 0)
                {
                    sb.AppendLine("");
                    sb.AppendLine("[RecentFiles]");
                }
                var i = 0;
                while (RecentFiles.Count > i)
                {
                    sb.AppendLine($"{(i + 1).ToString(CultureInfo.InvariantCulture)}=\"{RecentFiles[i]}\"");
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
                    LoadIni(new Model.AppIniFile(dlg.FileName, Encoding.UTF8));
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
                    LoadIni(new Model.AppIniFile(path, Encoding.UTF8));
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
