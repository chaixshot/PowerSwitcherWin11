using Petrroll.Helpers;
using PowerSwitcher.TrayApp.Configuration;
using PowerSwitcher.TrayApp.Services;
using PowerSwitcher.TrayApp.Windows;
using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Interop;

namespace PowerSwitcher.TrayApp
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern uint RegisterWindowMessage(string lpString);
        private static uint taskbarCreatedMessage;

        public HotKeyService HotKeyManager { get; private set; }
        public bool HotKeyFailed { get; private set; }

        public IPowerManager PowerManager { get; private set; }
        public TrayIcon TrayIcon { get; private set; }
        public ConfigurationInstance<PowerSwitcherSettings> Configuration { get; private set; }

        private Mutex _mMutex;
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            if (!tryToCreateMutex()) return;

            var configurationManager = new ConfigurationManagerXML<PowerSwitcherSettings>(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Petrroll", "PowerSwitcher", "PowerSwitcherSettings.xml"
                ));

            Configuration = new ConfigurationInstance<PowerSwitcherSettings>(configurationManager);
            migrateSettings();

            HotKeyManager = new HotKeyService();
            HotKeyFailed = false;

            PowerManager = new PowerManager();
            MainWindow = new MainWindow();

            TrayIcon = new TrayIcon(PowerManager, Configuration);
            TrayIcon.Show();

            Configuration.Data.PropertyChanged += Configuration_PropertyChanged;
            if (Configuration.Data.ShowOnShortcutSwitch) { registerHotkeyFromConfiguration(); }

            TrayIcon.CreateAltMenu();

            // Handle for explorer.exe restart
            {
                taskbarCreatedMessage = RegisterWindowMessage("TaskbarCreated");

                // MainWindow handle to listen for the OS broadcast
                var wih = new WindowInteropHelper(MainWindow);
                wih.EnsureHandle(); // Ensures the HWND exists even though the window is hidden

                HwndSource source = HwndSource.FromHwnd(wih.Handle);
                source.AddHook(HwndMessageHook);
            }
        }

        // The message loop handler
        private IntPtr HwndMessageHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            // Check if the message matches the "TaskbarCreated" ID assigned by Windows
            if (msg == taskbarCreatedMessage)
            {
                TrayIcon?.Close();
                TrayIcon = new TrayIcon(PowerManager, Configuration);
                TrayIcon.Show();
                TrayIcon.CreateAltMenu();
            }

            return IntPtr.Zero;
        }

        private void migrateSettings()
        {
            //Migration of shortcut because Creators update uses WinShift + S for screenshots
            if (Configuration.Data.ShowOnShortcutKey == System.Windows.Input.Key.S &&
                Configuration.Data.ShowOnShortcutKeyModifier == (KeyModifier.Shift | KeyModifier.Win))
            {
                Configuration.Data.ShowOnShortcutKey = System.Windows.Input.Key.L;
                Configuration.Save();
            }
        }

        private void Configuration_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(PowerSwitcherSettings.ShowOnShortcutSwitch))
            {
                if (Configuration.Data.ShowOnShortcutSwitch) { registerHotkeyFromConfiguration(); }
                else { unregisterHotkeyFromConfiguration(); }
            }
        }

        private void unregisterHotkeyFromConfiguration()
        {
            HotKeyManager.Unregister(new HotKey(Configuration.Data.ShowOnShortcutKey, Configuration.Data.ShowOnShortcutKeyModifier));
        }

        private bool registerHotkeyFromConfiguration()
        {
            var newHotKey = new HotKey(Configuration.Data.ShowOnShortcutKey, Configuration.Data.ShowOnShortcutKeyModifier);

            bool success = HotKeyManager.Register(newHotKey);
            if (!success) { HotKeyFailed = true; return false; }
            newHotKey.HotKeyFired += (this.MainWindow as MainWindow).ToggleWindowVisibility;

            return true;
        }

        private bool tryToCreateMutex()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var mutexName = string.Format(CultureInfo.InvariantCulture, "Local\\{{{0}}}{{{1}}}", assembly.GetType().GUID, assembly.GetName().Name);

            bool mutexCreated;

            _mMutex = new Mutex(true, mutexName, out mutexCreated);
            if (mutexCreated) { return true; }

            _mMutex = null;
            Current.Shutdown();
            return false;
        }

        private void DisposeMutex()
        {
            if (_mMutex == null) return;
            _mMutex.ReleaseMutex();
            _mMutex.Close();
            _mMutex = null;
        }

        private void App_OnExit(object sender, ExitEventArgs e)
        {
            DisposeMutex();
            PowerManager?.Dispose();
            HotKeyManager?.Dispose();
        }

        ~App()
        {
            DisposeMutex();
        }
    }
}
