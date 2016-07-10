using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace VbaSync {
    partial class App {
        void Application_Startup(object sender, StartupEventArgs e) {
            DispatcherUnhandledException += (s, e2) => {
                MessageBox.Show(e2.Exception.Message, VbaSync.Properties.Resources.MWTitle, MessageBoxButton.OK, MessageBoxImage.Error);
                e2.Handled = true;
            };
        }
    }
}
