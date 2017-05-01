using Ookii.Dialogs.Wpf;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
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
            DataContext = _vm = settings.Clone();
            _replaceSettings = replaceSettings;

            foreach (var cbi in LanguageComboBox.Items.Cast<ComboBoxItem>())
            {
                if (!string.IsNullOrEmpty(_vm.Language) && (string)cbi.Tag == _vm.Language)
                {
                    cbi.IsSelected = true;
                    break;
                }
                else if (string.IsNullOrEmpty(_vm.Language)
                    && Thread.CurrentThread.CurrentUICulture.ToString().StartsWith((string)cbi.Tag))
                {
                    cbi.IsSelected = true;
                    _vm.Language = (string)cbi.Tag;
                    break;
                }
            }

            _initialized = true;
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            FixQuotesEnclosingPath();
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

        private void FixQuotesEnclosingPath()
        {
            if (!string.IsNullOrEmpty(_vm.DiffTool) && _vm.DiffTool.Length > 2 && !File.Exists(_vm.DiffTool)
                && _vm.DiffTool.StartsWith("\"") && _vm.DiffTool.EndsWith("\"")
                && File.Exists(_vm.DiffTool.Substring(1, _vm.DiffTool.Length - 2)))
            {
                _vm.DiffTool = _vm.DiffTool.Substring(1, _vm.DiffTool.Length - 2);
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

        private void TextBoxFileDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
            {
                ((TextBox)sender).Text = files[0];
                BindingOperations.GetBindingExpression((TextBox)sender, TextBox.TextProperty)?.UpdateSource();
            }
        }

        private void TextBoxFilePreviewDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects |= DragDropEffects.Copy;
            }
            e.Handled = true;
        }

        private void TextBoxFilePreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }
    }
}
