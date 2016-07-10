using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Ookii.Dialogs.Wpf;

namespace VbaSync {
    partial class SettingsWindow {
        readonly Action<ISession> _replaceSession;

        public SettingsWindow(ISession session, Action<ISession> replaceSession) {
            InitializeComponent();
            DataContext = session.Copy();
            _replaceSession = replaceSession;
        }

        ISession Session => (ISession)DataContext;

        void ApplyButton_Click(object sender, RoutedEventArgs e) {
            _replaceSession(Session.Copy());
        }

        void CancelButton_Click(object sender, RoutedEventArgs e) {
            DialogResult = false;
            Close();
        }

        void DiffToolBrowseButton_Click(object sender, RoutedEventArgs e) {
            var dlg = new VistaOpenFileDialog {
                Filter = $"{Properties.Resources.SWOpenApplications}|*.exe",
                FilterIndex = 1
            };
            if (dlg.ShowDialog() == true) {
                Session.DiffTool = dlg.FileName;
            }
        }
        
        void OkButton_Click(object sender, RoutedEventArgs e) {
            ApplyButton_Click(null, null);
            DialogResult = true;
            Close();
        }
    }
}
