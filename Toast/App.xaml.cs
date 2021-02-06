using Microsoft.Toolkit.Uwp.Notifications;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace Toast
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            string shortcutPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Programs),
                "Toast.lnk");

            var execPath = Process.GetCurrentProcess().MainModule?.FileName;

            if (!File.Exists(shortcutPath))
            {
                ShortcutManager.RegisterAppForNotifications(shortcutPath, execPath, "", "Toast", "FBFEE6D6-8C90-44C2-9331-49300439D704");
            }

            // Register AUMID and COM server (for MSIX/sparse package apps, this no-ops)
            DesktopNotificationManagerCompat.RegisterAumidAndComServer<MyNotificationActivator>("Toast");
            // Register COM server and activator type
            DesktopNotificationManagerCompat.RegisterActivator<MyNotificationActivator>();

            base.OnStartup(e);
        }
    }
}
