using System;

namespace Toast
{
    public class ShortcutManager
    {
        public const string AppUserModelID = "Wpf-Toast-Notification-test";
        public const string ToastActivatorId = "FBFEE6D6-8C90-44C2-9331-49300439D704";

        public static void RegisterAppForNotifications(
            string shortcutPath,
            string appExecutablePath,
            string arguments,
            string appUserModelId,
            string activatorId)
        {
            using ShellLink link = new ShellLink()
            {
                TargetPath = appExecutablePath,
                Arguments = arguments,
                AppUserModelID = appUserModelId,
                ToastActivatorId = Guid.Parse(activatorId),
            };

            link.Save(shortcutPath);
        }
    }
}
