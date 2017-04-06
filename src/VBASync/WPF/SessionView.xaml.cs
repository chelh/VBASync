using Ookii.Dialogs.Wpf;
using System.Windows;
using System.Windows.Controls;
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

        public Control FocusControlOnEnter
        {
            get => (Control)GetValue(FocusControlOnEnterProperty);
            set => SetValue(FocusControlOnEnterProperty, value);
        }

        private Model.ISession Session => (Model.ISession)DataContext;

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
    }
}
