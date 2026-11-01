using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using Petrroll.Helpers;
using PowerSwitcher;
using PowerSwitcher.TrayApp;
using PowerSwitcher.TrayApp.Configuration;
using PowerSwitcher.TrayApp.Windows;
using Wpf.Ui.Tray.Controls;

internal static class Program
{
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern uint RegisterWindowMessage(string name);
    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

    [STAThread]
    private static int Main()
    {
        var app = new App { ShutdownMode = ShutdownMode.OnExplicitShutdown };
        try
        {
            // Initialize resources only: do not run Startup, acquire the production mutex,
            // read the user's settings, register hotkeys, or change real power plans.
            app.InitializeComponent();
            app.Startup -= (StartupEventHandler)Delegate.CreateDelegate(typeof(StartupEventHandler), app,
                typeof(App).GetMethod("Application_Startup", BindingFlags.Instance | BindingFlags.NonPublic));
            var power = new TestPowerManager();
            Set(app, "PowerManager", power);
            Set(app, "Configuration", new ConfigurationInstance<PowerSwitcherSettings>(new TestConfiguration()));
            app.MainWindow = new MainWindow();
            int initialSubscribers = power.SubscriberCount;
            var tray = new TrayIcon(power, app.Configuration);
            Set(app, "TrayIcon", tray);
            tray.Show();
            tray.CreateAltMenu();
            Pump();
            EnsureRegistered(tray);
            Check(RegisteredCount() == 1, "initial registration");

            uint message = RegisterWindowMessage("TaskbarCreated");
            typeof(App).GetField("taskbarCreatedMessage", BindingFlags.Static | BindingFlags.NonPublic).SetValue(null, message);
            var hook = (HwndSourceHook)Delegate.CreateDelegate(typeof(HwndSourceHook), app,
                typeof(App).GetMethod("HwndMessageHook", BindingFlags.Instance | BindingFlags.NonPublic));
            var hwnd = new WindowInteropHelper(app.MainWindow).Handle;
            var source = HwndSource.FromHwnd(hwnd);
            source.AddHook(hook);
            typeof(App).GetField("taskbarMessageSource", BindingFlags.Instance | BindingFlags.NonPublic)?.SetValue(app, source);
            try
            {
                // Target only this test window. No broadcast and no Explorer restart.
                for (int i = 1; i <= 12; i++)
                {
                    SendMessage(hwnd, message, IntPtr.Zero, IntPtr.Zero);
                    Pump();
                    EnsureRegistered(app.TrayIcon);
                    Check(RegisteredCount() == 1, "TaskbarCreated " + i + ": registered=" + RegisteredCount());
                    Check(power.SubscriberCount == initialSubscribers + 1, "only current tray subscribes to power events");
                }
                Console.WriteLine("PASS: 12 targeted TaskbarCreated messages keep exactly one registered icon and one tray subscriber.");
                var currentIcon = (NotifyIcon)app.TrayIcon.FindName("NotifyIcon");
                var previousImage = currentIcon.Icon;
                QueueThemeRefresh(app.TrayIcon);
                Pump();
                Check(!ReferenceEquals(previousImage, currentIcon.Icon), "background refresh reaches the UI thread");
                previousImage = currentIcon.Icon;
                QueueThemeRefresh(app.TrayIcon);
                app.TrayIcon.Close();
                Pump();
                Check(ReferenceEquals(previousImage, currentIcon.Icon), "queued refresh ignores the closed tray");
                Check(RegisteredCount() == 0, "closing the tray unregisters its icon");
                Check(power.SubscriberCount == initialSubscribers, "closing the tray removes its power subscriber");
                app.TrayIcon.UpdateTrayIcon();
                Pump();
                Check(RegisteredCount() == 0, "late theme refresh cannot restore a closed tray");
                Console.WriteLine("PASS: background theme refresh is dispatched; close removes icon/subscriber and ignores queued or late refresh.");

                tray = new TrayIcon(power, app.Configuration);
                Set(app, "TrayIcon", tray);
                tray.Show();
                tray.CreateAltMenu();
                Pump();
                EnsureRegistered(tray);
                bool powerDisposed = false;
                power.OnDispose = () =>
                {
                    Check(RegisteredCount() == 0, "tray is unregistered before power-manager disposal");
                    Check(power.SubscriberCount == initialSubscribers, "tray subscriber is removed before power-manager disposal");
                    powerDisposed = true;
                };
                typeof(App).GetMethod("App_OnExit", BindingFlags.Instance | BindingFlags.NonPublic)
                    .Invoke(app, new object[] { app, null });
                Pump();
                Check(powerDisposed, "application exit disposes the power manager");
                Check(RegisteredCount() == 0, "application exit unregisters the tray");
                Check(power.SubscriberCount == initialSubscribers, "application exit removes the tray subscriber");
                SendMessage(hwnd, message, IntPtr.Zero, IntPtr.Zero);
                Pump();
                Check(ReferenceEquals(tray, app.TrayIcon), "application exit removes the taskbar message hook");
                Check(RegisteredCount() == 0, "taskbar messages after exit cannot recreate an icon");
                Console.WriteLine("PASS: application exit removes the icon before disposing shared services and detaches the taskbar hook.");
            }
            finally { source.RemoveHook(hook); }
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine("FAIL: " + error); return 1; }
        finally
        {
            // Also clean up the deliberately failing baseline's leaked icons.
            var data = typeof(NotifyIcon).Assembly.GetType("Wpf.Ui.Tray.TrayData");
            var icons = (IEnumerable)data.GetProperty("NotifyIcons", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic).GetValue(null);
            foreach (var icon in icons.Cast<object>().ToArray())
                icon.GetType().GetMethod("Unregister").Invoke(icon, null);
            foreach (Window window in app.Windows.Cast<Window>().ToArray()) window.Close();
            app.Shutdown();
        }
    }

    private static void Set(App app, string property, object value) => typeof(App).GetProperty(property).SetValue(app, value);
    private static void Pump() => Dispatcher.CurrentDispatcher.Invoke(DispatcherPriority.ApplicationIdle, new Action(() => { }));
    private static void Check(bool value, string message) { if (!value) throw new Exception(message); }
    private static void QueueThemeRefresh(TrayIcon tray)
    {
        Exception failure = null;
        var worker = new Thread(() => { try { tray.UpdateTrayIcon(); } catch (Exception error) { failure = error; } });
        worker.Start();
        worker.Join();
        Check(failure == null, "background theme refresh: " + failure);
    }
    private static void EnsureRegistered(TrayIcon tray)
    {
        var icon = (NotifyIcon)tray.FindName("NotifyIcon");
        if (!icon.IsRegistered) icon.Register();
        Check(icon.IsRegistered, "test icon registered");
    }
    private static int RegisteredCount()
    {
        var data = typeof(NotifyIcon).Assembly.GetType("Wpf.Ui.Tray.TrayData");
        var icons = (IEnumerable)data.GetProperty("NotifyIcons", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic).GetValue(null);
        return icons.Cast<object>().Count(icon => (bool)icon.GetType().GetProperty("IsRegistered").GetValue(icon));
    }
    private sealed class TestConfiguration : IConfigurationManger<PowerSwitcherSettings>
    {
        public PowerSwitcherSettings DeserializeOrDefault() => new PowerSwitcherSettings { AutomaticOnACSwitch = false, ShowOnShortcutSwitch = false };
        public void SerializeConfiguration(PowerSwitcherSettings value) { throw new Exception("Unexpected settings write"); }
    }
    private sealed class TestPowerManager : IPowerManager
    {
        private PropertyChangedEventHandler changed;
        public event PropertyChangedEventHandler PropertyChanged { add { changed += value; } remove { changed -= value; } }
        public int SubscriberCount => changed == null ? 0 : changed.GetInvocationList().Length;
        public Action OnDispose { get; set; }
        public ObservableCollection<IPowerSchema> Schemas { get; } = new ObservableCollection<IPowerSchema>();
        public IPowerSchema CurrentSchema => null;
        public PowerPlugStatus CurrentPowerStatus => PowerPlugStatus.Online;
        public void UpdateSchemas() { }
        public void SetPowerSchema(IPowerSchema schema) { throw new Exception("Unexpected power change"); }
        public void SetPowerSchema(Guid schema) { throw new Exception("Unexpected power change"); }
        public void Dispose() { OnDispose?.Invoke(); }
    }
}
