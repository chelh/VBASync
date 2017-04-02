using Ookii.Dialogs.Wpf;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using VBASync.Localization;
using VBASync.Model;

namespace VBASync.WPF
{
    internal partial class SettingsWindow
    {
        private readonly bool _initialized;
        private readonly Action<ISession> _replaceSession;

        public SettingsWindow(ISession session, Action<ISession> replaceSession)
        {
            InitializeComponent();
            DataContext = session.Copy();
            _replaceSession = replaceSession;
            foreach (var cbi in LanguageComboBox.Items.Cast<ComboBoxItem>())
            {
                if ((string)cbi.Tag == Session.Language)
                {
                    cbi.IsSelected = true;
                }
            }
            _initialized = true;
        }

        private ISession Session => (ISession)DataContext;

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            _replaceSession(Session.Copy());
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void DiffToolBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new VistaOpenFileDialog {
                Filter = $"{VBASyncResources.SWOpenApplications}|*.exe",
                FilterIndex = 1
            };
            if (dlg.ShowDialog() == true)
            {
                Session.DiffTool = dlg.FileName;
            }
        }

        private void LanguageComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_initialized)
            {
                return;
            }
            Session.Language = (e.AddedItems[0] as ComboBoxItem).Tag as string;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            ApplyButton_Click(null, null);
            DialogResult = true;
            Close();
        }
    }
}
