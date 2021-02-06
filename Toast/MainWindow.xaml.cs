using Microsoft.Toolkit.Uwp.Notifications;
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
            // Construct the visuals of the toast (using Notifications library)
            ToastContent toastContent = new ToastContentBuilder()
                .AddToastActivationInfo("action=viewConversation&conversationId=5", ToastActivationType.Foreground)
                .AddText("Hello world!")
                .GetToastContent();

            // And create the toast notification
            var toast = new ToastNotification(toastContent.GetXml());

            // And then show it
            DesktopNotificationManagerCompat.CreateToastNotifier().Show(toast);
        }
    }
}
