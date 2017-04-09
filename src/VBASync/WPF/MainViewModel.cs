using System;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows;

namespace VBASync.WPF
{
    internal class MainViewModel : ViewModelBase
    {
        private SessionViewModel _session;
        private SettingsViewModel _settings;

        internal MainViewModel()
        {
            OpenRecentCommand = new WpfCommand(s => LoadRecent(s));
            RecentFiles = new BindingList<string>();
            // promote ListChanged events to PropertyChanged, so the
            // recent files list in the main menu stays up to date
            RecentFiles.ListChanged += (s, e) => OnPropertyChanged(nameof(RecentFiles));
        }

        public WpfCommand OpenRecentCommand { get; }
        public BindingList<string> RecentFiles { get; }

        public SessionViewModel Session
        {
            get => _session;
            set => SetField(ref _session, value, nameof(Session));
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
                    if (!File.Exists(RecentFiles[i]))
                    {
                        throw new FileNotFoundException();
                    }
                    LoadIni(new Model.AppIniFile(RecentFiles[i], Encoding.UTF8));
                    AddRecentFile(RecentFiles[i]);
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
                MessageBox.Show(ex.Message, Localization.VBASyncResources.MWTitle,
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
