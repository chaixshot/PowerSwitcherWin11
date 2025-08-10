using Microsoft.Win32;
using System.Windows;
using System.Windows.Media;

namespace PowerSwitcher.TrayApp.Services
{
    ////
    //  Code heavily inspired by https://github.com/File-New-Project/EarTrumpet/blob/master/EarTrumpet/Services/ThemeService.cs
    ////
    public class ThemeService
    {
        public static bool IsWindowTransparencyEnabled
        {
            get { return !SystemParameters.HighContrast && UserSystemPreferencesService.IsTransparencyEnabled; }
        }

        public static bool IsUseLightTheme
        {
            get { return Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize").GetValue("AppsUseLightTheme").ToString() == "1"; }
        }

        public static void UpdateThemeResources(ResourceDictionary dictionary)
        {
            dictionary["WindowBackground"] = new SolidColorBrush(GetWindowBackgroundColor());

            if (IsUseLightTheme)
            {
                ReplaceBrush(dictionary, "WindowForeground", "ImmersiveApplicationTextLightTheme");
            }
            else
            {
                ReplaceBrush(dictionary, "WindowForeground", "ImmersiveApplicationTextDarkTheme");
            }
            ReplaceBrushWithOpacity(dictionary, "SelectedItemBackground", "ImmersiveSystemAccent", 0.9);
            ReplaceBrushWithOpacity(dictionary, "MouseOverSelectedItemBackground", "ImmersiveSystemAccent", 1.0);
            ReplaceBrushWithOpacity(dictionary, "MouseOverItemBackground", "ImmersiveControlLightSelectHighlightSelectedHover", 0.3);
            (((App)Application.Current).TrayApp as TrayApp).UpdateTrayIcon();
        }

        private static Color GetWindowBackgroundColor()
        {
            string resource;
            if (IsUseLightTheme)
            {
                resource = "ImmersiveApplicationTextDarkTheme";
            }
            else
            {
                resource = "ImmersiveApplicationTextLightTheme";
            }

            var color = AccentColorService.GetColorByTypeName(resource);
            color.A = (byte)(IsWindowTransparencyEnabled ? 190 : 255);
            return color;
        }

        private static void SetBrushWithOpacity(ResourceDictionary dictionary, string name, string immersiveAccentName, double opacity)
        {
            var color = AccentColorService.GetColorByTypeName(immersiveAccentName);
            color.A = (byte)(opacity * 255);
            ((SolidColorBrush)dictionary[name]).Color = color;
        }

        private static void SetBrush(ResourceDictionary dictionary, string name, string immersiveAccentName)
        {
            SetBrushWithOpacity(dictionary, name, immersiveAccentName, 1.0);
        }

        private static void ReplaceBrush(ResourceDictionary dictionary, string name, string immersiveAccentName)
        {
            dictionary[name] = new SolidColorBrush(AccentColorService.GetColorByTypeName(immersiveAccentName));
        }
        private static void ReplaceBrushWithOpacity(ResourceDictionary dictionary, string name, string immersiveAccentName, double opacity)
        {
            var color = AccentColorService.GetColorByTypeName(immersiveAccentName);
            color.A = (byte)(opacity * 255);
            dictionary[name] = new SolidColorBrush(color);
        }
    }
}
