using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace VBASync.WPF
{
    internal partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            InitializeComponent();

            var asm = Assembly.GetExecutingAssembly();
            var version = asm.GetName().Version;

            VersionLabel.Content = ((string)VersionLabel.Content).Replace("{0}",
                MainWindow.Version.ToString());
            CopyrightLabel.Content = ((string)CopyrightLabel.Content).Replace("{0}",
                MainWindow.CopyrightYear.ToString());
        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
