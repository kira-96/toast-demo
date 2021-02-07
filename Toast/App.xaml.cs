using Microsoft.Toolkit.Uwp.Notifications;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

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
                // Create a shortcut
                ShortcutManager.RegisterAppForNotifications(shortcutPath, execPath, null, ShortcutManager.AppUserModelID, ShortcutManager.ToastActivatorId);
            }

            // Register AUMID and COM server (for MSIX/sparse package apps, this no-ops)
            DesktopNotificationManagerCompat.RegisterAumidAndComServer<MyNotificationActivator>(ShortcutManager.AppUserModelID);
            // Register COM server and activator type
            DesktopNotificationManagerCompat.RegisterActivator<MyNotificationActivator>();

            // If launched from a toast
            if (e.Args.Contains("-ToastActivated"))
            {
                // Our NotificationActivator code will run after this completes,
                // and will show a window if necessary.
            }
            else
            {
                // Show the window
                // In App.xaml, be sure to remove the StartupUri so that a window doesn't
                // get created by default, since we're creating windows ourselves (and sometimes we
                // don't want to create a window if handling a background activation).
                new MainWindow().Show();
            }

            base.OnStartup(e);

            DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
        }

        private void CurrentDomain_UnhandledException(object s, UnhandledExceptionEventArgs e)
        {
            if (e.IsTerminating)
            {
                Current.Shutdown(-1);
            }
            else
            {
                MessageBox.Show(((Exception)e.ExceptionObject).Message, "An error occurred!");
            }
        }

        private void TaskScheduler_UnobservedTaskException(object s, UnobservedTaskExceptionEventArgs e)
        {
            e.SetObserved();

            MessageBox.Show(e.Exception.Message, "An error occurred!");
        }

        private void App_DispatcherUnhandledException(object s, DispatcherUnhandledExceptionEventArgs e)
        {
            e.Handled = true;

            MessageBox.Show(e.Exception.Message, "An error occurred!");
        }
    }
}
