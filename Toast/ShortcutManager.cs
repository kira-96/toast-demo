using System;

namespace Toast
{
    public class ShortcutManager
    {
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
                AppUserModelID = appUserModelId,
                ToastActivatorId = Guid.Parse(activatorId),
                Arguments = arguments
            };

            link.Save(shortcutPath);
        }
    }
}
