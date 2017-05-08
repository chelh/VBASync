using Ookii.Dialogs.Wpf;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using VBASync.Localization;

namespace VBASync.WPF
{
    internal partial class SessionView : UserControl
    {
        public static readonly DependencyProperty FocusControlOnEnterProperty
            = DependencyProperty.Register("FocusControlOnEnter", typeof(Control), typeof(SessionView));

        public SessionView()
        {
            InitializeComponent();
        }

        public bool DataValidationFaulted => !File.Exists(FaultedFilePath) || !Directory.Exists(FaultedFolderPath);
        public string FaultedFilePath => FileBrowseBox.Text;
        public string FaultedFolderPath => FolderBrowseBox.Text;

        public Control FocusControlOnEnter
        {
            get => (Control)GetValue(FocusControlOnEnterProperty);
            set => SetValue(FocusControlOnEnterProperty, value);
        }

        private SessionViewModel Session => (SessionViewModel)DataContext;

        private void FileBrowseBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                FocusControlOnEnter.Focus();
            }
        }

        private void FileBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new VistaOpenFileDialog
            {
                Filter = $"{VBASyncResources.MWOpenAllFiles}|*.*|"
                    + $"{VBASyncResources.MWOpenAllSupported}|*.doc;*.dot;*.xls;*.xlt;*.docm;*.dotm;*.docb;*.xlsm;*.xla;*.xlam;*.xlsb;"
                    + "*.pptm;*.potm;*.ppam;*.ppsm;*.sldm;*.docx;*.dotx;*.xlsx;*.xltx;*.pptx;*.potx;*.ppsx;*.sldx;*.otm;*.bin|"
                    + $"{VBASyncResources.MWOpenWord97}|*.doc;*.dot|"
                    + $"{VBASyncResources.MWOpenExcel97}|*.xls;*.xlt;*.xla|"
                    + $"{VBASyncResources.MWOpenWord07}|*.docx;*.docm;*.dotx;*.dotm;*.docb|"
                    + $"{VBASyncResources.MWOpenExcel07}|*.xlsx;*.xlsm;*.xltx;*.xltm;*.xlsb;*.xlam|"
                    + $"{VBASyncResources.MWOpenPpt07}|*.pptx;*.pptm;*.potx;*.potm;*.ppam;*.ppsx;*.ppsm;*.sldx;*.sldm|"
                    + $"{VBASyncResources.MWOpenOutlook}|*.otm|"
                    + $"{VBASyncResources.MWOpenSAlone}|*.bin",
                FilterIndex = 2
            };
            if (dlg.ShowDialog() == true)
            {
                Session.FilePath = dlg.FileName;
            }
        }

        private void FolderBrowseBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                FocusControlOnEnter.Focus();
            }
        }

        private void FolderBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new VistaFolderBrowserDialog();
            if (dlg.ShowDialog() == true)
            {
                Session.FolderPath = dlg.SelectedPath;
            }
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
