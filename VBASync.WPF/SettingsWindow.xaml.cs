using Ookii.Dialogs.Wpf;
using System;
using System.Windows;

namespace VbaSync {
    internal partial class SettingsWindow {
        private readonly Action<ISession> _replaceSession;

        public SettingsWindow(ISession session, Action<ISession> replaceSession) {
            InitializeComponent();
            DataContext = session.Copy();
            _replaceSession = replaceSession;
        }

        private ISession Session => (ISession)DataContext;

        private void ApplyButton_Click(object sender, RoutedEventArgs e) {
            _replaceSession(Session.Copy());
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e) {
            DialogResult = false;
            Close();
        }

        private void DiffToolBrowseButton_Click(object sender, RoutedEventArgs e) {
            var dlg = new VistaOpenFileDialog {
                Filter = $"{Properties.Resources.SWOpenApplications}|*.exe",
                FilterIndex = 1
            };
            if (dlg.ShowDialog() == true) {
                Session.DiffTool = dlg.FileName;
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e) {
            ApplyButton_Click(null, null);
            DialogResult = true;
            Close();
        }
    }
}
