using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Documents;
using VBASync.Localization;

namespace VBASync.WPF
{
    internal partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            InitializeComponent();

            VersionLabel.Content = ((string)VersionLabel.Content).Replace("{0}",
                MainWindow.Version.ToString());

            var exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var copyrightText = VBASyncResources.AWCopyright
                .Replace("{0}", MainWindow.CopyrightYear.ToString());
            CopyrightBlock.Inlines.Clear();
            CopyrightBlock.Inlines.Add(TextBefore("{1}", copyrightText));
            var licenseRtfHyperlink = new Hyperlink
            {
                NavigateUri = new Uri("file:///")
            };
            licenseRtfHyperlink.Inlines.Add("LICENSE.rtf");
            licenseRtfHyperlink.RequestNavigate += (s, e)
                => Process.Start(Path.Combine(exeDir, "LICENSE.rtf"));
            CopyrightBlock.Inlines.Add(licenseRtfHyperlink);
            CopyrightBlock.Inlines.Add(TextAfter("{1}", copyrightText));

            WebsiteBlock.Inlines.Clear();
            WebsiteBlock.Inlines.Add(TextBefore("{0}", VBASyncResources.AWWebsite));
            var websiteHyperlink = new Hyperlink
            {
                NavigateUri = new Uri(MainWindow.SupportUrl)
            };
            websiteHyperlink.Inlines.Add(MainWindow.SupportUrl);
            websiteHyperlink.RequestNavigate += (s, e) => Process.Start(e.Uri.ToString());
            WebsiteBlock.Inlines.Add(websiteHyperlink);
            WebsiteBlock.Inlines.Add(TextAfter("{0}", VBASyncResources.AWWebsite));
        }

        private static string TextAfter(string separator, string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }
            var idx = text.IndexOf(separator);
            if (idx < 0)
            {
                return "";
            }
            return text.Substring(idx + separator.Length, text.Length - idx - separator.Length);
        }

        private static string TextBefore(string separator, string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }
            var idx = text.IndexOf(separator);
            if (idx <= 0)
            {
                return "";
            }
            return text.Substring(0, idx);
        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
