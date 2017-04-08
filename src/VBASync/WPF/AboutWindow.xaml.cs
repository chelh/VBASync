using System.Windows;

namespace VBASync.WPF
{
    internal partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            InitializeComponent();

            VersionLabel.Content = ((string)VersionLabel.Content).Replace("{0}",
                MainWindow.Version.ToString());
            CopyrightLabel.Content = ((string)CopyrightLabel.Content).Replace("{0}",
                MainWindow.CopyrightYear.ToString());
            WebsiteLabel.Content = ((string)WebsiteLabel.Content).Replace("{0}",
                MainWindow.SupportUrl);
        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
