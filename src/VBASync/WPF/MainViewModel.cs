using System.ComponentModel;
using System.IO;
using System.Text;

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

        public void AddRecentFile(string relativePath)
        {
            var absPath = Path.GetFullPath(relativePath);
            if (RecentFiles.Contains(absPath))
            {
                RecentFiles.Remove(absPath);
            }
            RecentFiles.Insert(0, absPath);
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
            var i = int.Parse(index) - 1;
            LoadIni(new Model.AppIniFile(RecentFiles[i], Encoding.UTF8));
            AddRecentFile(RecentFiles[i]);
        }
    }
}
