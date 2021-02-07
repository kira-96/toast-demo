using Microsoft.Toolkit.Uwp.Notifications;
using System;
using System.Windows;
using Windows.UI.Notifications;

namespace Toast
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ToastButton_Click(object s, RoutedEventArgs e)
        {
            if (!CheckNotificationSupported)
            {
                throw new NotSupportedException("Windows version must be higher than 10.0.17763.0");
            }

            // Construct the visuals of the toast (using Notifications library)
            ToastContent toastContent = new ToastContentBuilder()
                .SetToastScenario(ToastScenario.Reminder)
                .AddToastActivationInfo("action=viewEvent&eventId=1983", ToastActivationType.Foreground)
                .AddText("Adaptive Tiles Meeting")
                .AddText("Conf Room 2001 / Building 135")
                .AddText("10:00 AM - 10:30 AM")
                .AddComboBox("snoozeTime", "15", ("1", "1 minute"),
                                                 ("15", "15 minutes"),
                                                 ("60", "1 hour"),
                                                 ("240", "4 hours"),
                                                 ("1440", "1 day"))
                .AddButton(new ToastButtonSnooze() { SelectionBoxId = "snoozeTime" })
                .AddButton(new ToastButtonDismiss())
                .GetToastContent();

            // And create the toast notification
            var toast = new ToastNotification(toastContent.GetXml());

            // And then show it
            DesktopNotificationManagerCompat.CreateToastNotifier().Show(toast);
        }

        private void ClearButton_Click(object s, RoutedEventArgs e)
        {
            if (!CheckNotificationSupported)
            {
                throw new NotSupportedException("Windows version must be higher than 10.0.17763.0");
            }

            DesktopNotificationManagerCompat.History.Clear();
        }

        private bool CheckNotificationSupported => Environment.OSVersion.Version >= Version.Parse("10.0.17763.0");
    }
}
