using Ookii.Dialogs.Wpf;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Input;
using VBASync.Localization;
using VBASync.Model;

namespace VBASync.WPF {
    internal partial class MainWindow {
        internal const int CopyrightYear = 2017;
        internal const string SupportUrl = "https://github.com/chelh/VBASync";

        internal static readonly Version Version = new Version(2, 1, 0);

        private readonly MainViewModel _vm;

        private bool _doUpdateIncludeAll = true;

        public MainWindow(Startup startup) {
            InitializeComponent();

            DataContext = _vm = new MainViewModel(startup);
            DataContextChanged += (s, e) => QuietRefreshIfInputsOk();
            _vm.Session.PropertyChanged += (s, e) => QuietRefreshIfInputsOk();
            QuietRefreshIfInputsOk();
        }

        private ISession Session => _vm.Session;

        internal void SettingsMenu_Click(object sender, RoutedEventArgs e)
        {
            new SettingsWindow(_vm.Settings, s => _vm.Settings = s).ShowDialog();
        }

        private void AboutMenu_Click(object sender, RoutedEventArgs e)
        {
            new AboutWindow().ShowDialog();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            CheckAndFixErrors();

            var changes = ChangesGrid.DataContext as ChangesViewModel;
            var committedChanges = changes?.Where(p => p.Commit).ToList();
            if (committedChanges == null || committedChanges.Count == 0)
            {
                return;
            }

            _vm.ActiveSession.Apply(committedChanges);

            QuietRefreshIfInputsOk();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e) {
            Application.Current.Shutdown();
        }

        private void ChangesGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e) {
            var sel = (Patch)ChangesGrid.SelectedItem;
            var fileName = sel.ModuleName + (sel.ChangeType != ChangeType.ChangeFormControls ? ModuleProcessing.ExtensionFromType(sel.ModuleType) : ".frx");
            string oldPath;
            string newPath;
            if (Session.Action == ActionType.Extract) {
                oldPath = Path.Combine(Session.FolderPath, fileName);
                newPath = Path.Combine(_vm.ActiveSession.FolderPath, fileName);
            } else {
                oldPath = Path.Combine(_vm.ActiveSession.FolderPath, fileName);
                newPath = Path.Combine(Session.FolderPath, fileName);
            }
            if (sel.ChangeType == ChangeType.ChangeFormControls) {
                Lib.FrxFilesAreDifferent(oldPath, newPath, out var explain);
                MessageBox.Show(explain, VBASyncResources.ExplainFrxTitle, MessageBoxButton.OK, MessageBoxImage.Information);
            } else {
                var diffExePath = Environment.ExpandEnvironmentVariables(_vm.Settings.DiffTool);
                if (!File.Exists(oldPath) || !File.Exists(newPath) || !File.Exists(diffExePath)) {
                    return;
                }
                var p = new Process {
                    StartInfo = new ProcessStartInfo(diffExePath, _vm.Settings.DiffToolParameters.Replace("{OldFile}", oldPath).Replace("{NewFile}", newPath)) {
                        UseShellExecute = false
                    }
                };
                p.Start();
            }
        }

        private void CheckAndFixErrors()
        {
            var filePath = _vm.Session.FilePath;
            var folderPath = _vm.Session.FolderPath;
            if (!string.IsNullOrEmpty(filePath) && filePath.Length > 2 && !FileOrFolderExists(filePath)
                && filePath.StartsWith("\"") && filePath.EndsWith("\"")
                && FileOrFolderExistsOrIsRooted(filePath.Substring(1, filePath.Length - 2)))
            {
                filePath = filePath.Substring(1, filePath.Length - 2);
                _vm.Session.FilePath = filePath;
            }
            if (!string.IsNullOrEmpty(folderPath) && folderPath.Length > 2 && !FileOrFolderExists(folderPath)
                && folderPath.StartsWith("\"") && folderPath.EndsWith("\"")
                && FileOrFolderExistsOrIsRooted(folderPath.Substring(1, folderPath.Length - 2)))
            {
                folderPath = folderPath.Substring(1, folderPath.Length - 2);
                _vm.Session.FolderPath = folderPath;
            }
            if (!File.Exists(filePath) && !Directory.Exists(folderPath) && File.Exists(folderPath) && Directory.Exists(filePath))
            {
                _vm.Session.FolderPath = filePath;
                _vm.Session.FilePath = folderPath;
            }

            bool FileOrFolderExists(string path) => File.Exists(path) || Directory.Exists(path);
            bool FileOrFolderExistsOrIsRooted(string path) => FileOrFolderExists(path) || Path.IsPathRooted(path);
        }

        private void ExitMenu_Click(object sender, RoutedEventArgs e) {
            CancelButton_Click(null, null);
        }

        private void IncludeAllBox_Click(object sender, RoutedEventArgs e)
        {
            var vm = ChangesGrid.DataContext as ChangesViewModel;
            if (vm == null || IncludeAllBox.IsChecked == null) {
                return;
            }
            try {
                _doUpdateIncludeAll = false;
                foreach (var ch in vm) {
                    ch.Commit = IncludeAllBox.IsChecked.Value;
                }
            } finally {
                _doUpdateIncludeAll = true;
            }
            ChangesGrid.Items.Refresh();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e) {
            ApplyButton_Click(null, null);
            Application.Current.Shutdown();
        }

        private void QuietRefreshIfInputsOk() {
            CheckAndFixErrors();
            if (!File.Exists(Session.FilePath) || !Directory.Exists(Session.FolderPath)) {
                return;
            }
            try {
                RefreshButton_Click(null, null);
            } catch {
                ChangesGrid.DataContext = null;
            }
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e) {
            if (string.IsNullOrEmpty(Session.FolderPath) || string.IsNullOrEmpty(Session.FilePath)) {
                return;
            }
            CheckAndFixErrors();
            _vm.RefreshActiveSession();

            var changes = new ChangesViewModel(_vm.ActiveSession.GetPatches());
            ChangesGrid.DataContext = changes;
            foreach (var p in changes) {
                p.CommitChanged += (s2, e2) => UpdateIncludeAllBox();
            }
            UpdateIncludeAllBox();
        }

        private void UpdateIncludeAllBox() {
            if (!_doUpdateIncludeAll) {
                return;
            }
            var vm = (ChangesViewModel)ChangesGrid.DataContext;
            if (vm.All(p => p.Commit)) {
                IncludeAllBox.IsChecked = true;
            } else if (vm.All(p => !p.Commit)) {
                IncludeAllBox.IsChecked = false;
            } else {
                IncludeAllBox.IsChecked = null;
            }
        }

        private void Window_Closed(object sender, EventArgs e) {
            _vm.ActiveSession?.Dispose();

            string lastSessionPath;
            if (_vm.Settings.Portable)
            {
                var exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                lastSessionPath = Path.Combine(exeDir, "LastSession.ini");
            }
            else
            {
                lastSessionPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "VBA Sync Tool", "LastSession.ini");
                Directory.CreateDirectory(Path.GetDirectoryName(lastSessionPath));
            }
            using (var st = new FileStream(lastSessionPath, FileMode.Create))
            {
                _vm.SaveSession(st, true);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e) {
            QuietRefreshIfInputsOk();

            if (Session.AutoRun) {
                OkButton_Click(null, null);
            }
        }
    }
}
