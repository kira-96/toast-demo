using Microsoft.QueryStringDotNET;
using Microsoft.Toolkit.Uwp.Notifications;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;

namespace Toast
{
    // The GUID CLSID must be unique to your app. Create a new GUID if copying this code.
    [ClassInterface(ClassInterfaceType.None)]
#pragma warning disable CS0618
    [ComSourceInterfaces(typeof(INotificationActivationCallback))]
#pragma warning restore CS0618
    [Guid(ShortcutManager.ToastActivatorId), ComVisible(true)]
    public class MyNotificationActivator : NotificationActivator
    {
        public override void OnActivated(string arguments, NotificationUserInput userInput, string appUserModelId)
        {
            // TODO: Handle activation
            // The OnActivated method is not called on the UI thread. 
            // If you'd like to perform UI thread operations, you must call 
            // Application.Current.Dispatcher.Invoke(callback).
            Application.Current.Dispatcher.Invoke(delegate
            {
                // Tapping on the top-level header launches with empty args
                if (arguments.Length == 0)
                {
                    // Perform a normal launch
                    OpenWindowIfNeeded();
                    return;
                }

                // Parse the query string (using NuGet package QueryString.NET)
                QueryString args = QueryString.Parse(arguments);

                // See what action is being requested
                switch (args["action"])
                {
                    case "viewEvent":
                        OpenWindowIfNeeded();
                        MessageBox.Show($"EventID={args["eventId"]}", appUserModelId);
                        break;
                    default:
                        break;
                }
            });
        }

        private void OpenWindowIfNeeded()
        {
            // Make sure we have a window open (in case user clicked toast while app closed)
            if (Application.Current.Windows.Count == 0)
            {
                new MainWindow().Show();
            }

            // Activate the window, bringing it to focus
            Application.Current.Windows[0].Activate();

            // And make sure to maximize the window too, in case it was currently minimized
            Application.Current.Windows[0].WindowState = WindowState.Normal;
        }
    }
}
