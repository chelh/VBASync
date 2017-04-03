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
        private readonly Action<SettingsViewModel> _replaceSettings;
        private readonly SettingsViewModel _vm;

        public SettingsWindow(SettingsViewModel settings, Action<SettingsViewModel> replaceSettings)
        {
            InitializeComponent();
            _vm = settings.Clone();
            DataContext = _vm;
            _replaceSettings = replaceSettings;
            foreach (var cbi in LanguageComboBox.Items.Cast<ComboBoxItem>())
            {
                if ((string)cbi.Tag == _vm.Language)
                {
                    cbi.IsSelected = true;
                }
            }
            _initialized = true;
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            _replaceSettings(_vm.Clone());
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
                _vm.DiffTool = dlg.FileName;
            }
        }

        private void LanguageComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_initialized)
            {
                return;
            }
            _vm.Language = (e.AddedItems[0] as ComboBoxItem).Tag as string;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            ApplyButton_Click(null, null);
            DialogResult = true;
            Close();
        }
    }
}
