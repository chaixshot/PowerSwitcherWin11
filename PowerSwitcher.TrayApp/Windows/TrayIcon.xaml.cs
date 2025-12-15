using Microsoft.Win32;
using Petrroll.Helpers;
using PowerSwitcher.TrayApp.Configuration;
using PowerSwitcher.TrayApp.Resources;
using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media.Imaging;
using Wpf.Ui.Controls;

namespace PowerSwitcher.TrayApp.Windows
{
    /// <summary>
    /// Interaction logic for Trayicon.xaml
    /// </summary>
    public partial class TrayIcon : Window
    {
        public event Action ShowFlyout;
        IPowerManager pwrManager;
        ConfigurationInstance<PowerSwitcherSettings> configuration;
        [DllImport("UXTheme.dll", SetLastError = true, EntryPoint = "#138")]
        public static extern bool ShouldSystemUseDarkMode();

        public TrayIcon(IPowerManager powerManager, ConfigurationInstance<PowerSwitcherSettings> config)
        {
            InitializeComponent();

            Loaded += (s, e) =>
            {
                this.Hide();
            };

            this.pwrManager = powerManager;
            pwrManager.PropertyChanged += PwrManager_PropertyChanged;
            configuration = config;

            this.ShowFlyout += (((App)Application.Current).MainWindow as MainWindow).ToggleWindowVisibility;

            //Run automatic on-off-AC change at boot
            powerStatusChanged();

            // Tray light/dark theme
            UpdateTrayIcon();
            SystemEvents.UserPreferenceChanged += (s, e) => UpdateTrayIcon();
        }

        public void UpdateTrayIcon()
        {
            string icon = "pack://application:,,,/PowerSwitcher.TrayApp;component/Tray_Light.ico";
            if (ShouldSystemUseDarkMode())
                icon = "pack://application:,,,/PowerSwitcher.TrayApp;component/Tray_Dark.ico";
            NotifyIcon.Icon = new BitmapImage(new Uri(icon));
        }

        public void CreateAltMenu()
        {
            NotifyIcon.RightClick += ContextMenu_Popup;

            MenuItem settingsOffACItem = new MenuItem() { Header = AppStrings.SchemaToSwitchOffAc };
            SettingsMenuItem.Items.Insert(0, settingsOffACItem);

            MenuItem settingsOnAC = new MenuItem() { Header = AppStrings.SchemaToSwitchOnAc };
            SettingsMenuItem.Items.Insert(1, settingsOnAC);

            var AutomaticOnOffACSwitch = new MenuItem()
            {
                Header = AppStrings.AutomaticOnOffACSwitch,
                Icon = new SymbolIcon(configuration.Data.AutomaticOnACSwitch ? SymbolRegular.Checkmark24 : SymbolRegular.Empty),
            };
            AutomaticOnOffACSwitch.Click += AutomaticSwitchItem_Click;
            SettingsMenuItem.Items.Add(AutomaticOnOffACSwitch);

            var HideFlyoutAfterSchemaChangeSwitch = new MenuItem()
            {
                Header = AppStrings.HideFlyoutAfterSchemaChangeSwitch,
                Icon = new SymbolIcon(configuration.Data.AutomaticFlyoutHideAfterClick ? SymbolRegular.Checkmark24 : SymbolRegular.Empty),
            };
            HideFlyoutAfterSchemaChangeSwitch.Click += AutomaticHideItem_Click;
            SettingsMenuItem.Items.Add(HideFlyoutAfterSchemaChangeSwitch);

            var ShowOnlyDefaultSchemas = new MenuItem()
            {
                Header = AppStrings.ShowOnlyDefaultSchemas,
                Icon = new SymbolIcon(configuration.Data.ShowOnlyDefaultSchemas ? SymbolRegular.Checkmark24 : SymbolRegular.Empty),
            };
            ShowOnlyDefaultSchemas.Click += OnlyDefaultSchemas_Click;
            SettingsMenuItem.Items.Add(ShowOnlyDefaultSchemas);

            var ToggleOnShowrtcutSwitch = new MenuItem()
            {
                Header = $"{AppStrings.ToggleOnShowrtcutSwitch} ({configuration.Data.ShowOnShortcutKeyModifier} + {configuration.Data.ShowOnShortcutKey})",
                IsEnabled = !(Application.Current as App).HotKeyFailed,
                Icon = new SymbolIcon(configuration.Data.ShowOnShortcutSwitch ? SymbolRegular.Checkmark24 : SymbolRegular.Empty),
            };
            ToggleOnShowrtcutSwitch.Click += EnableShortcutsToggleItem_Click;
            SettingsMenuItem.Items.Add(ToggleOnShowrtcutSwitch);
        }

        #region ContextMenuItemRelatedStuff
        private void ContextMenu_Popup(object sender, EventArgs e)
        {
            clearPowerSchemasInTray();

            pwrManager.UpdateSchemas();
            foreach (var powerSchema in pwrManager.Schemas)
            {
                updateTrayMenuWithPowerSchema(powerSchema);
            }
        }
        private void updateTrayMenuWithPowerSchema(IPowerSchema powerSchema)
        {
            var newItemMain = getNewPowerSchemaItem(
                powerSchema,
                (s, ea) => switchToPowerSchema(powerSchema),
                powerSchema.IsActive
                );
            NotifyIcon.Menu.Items.Insert(0, newItemMain);

            //ItemSettingsOffAC
            var newItemSettingsOffAC = getNewPowerSchemaItem(
                powerSchema,
                (s, ea) => setPowerSchemaAsOffAC(powerSchema),
                (powerSchema.Guid == configuration.Data.AutomaticPlanGuidOffAC)
                );
            MenuItem settingsOffAC = SettingsMenuItem.Items.GetItemAt(0) as MenuItem;
            settingsOffAC.Items.Add(newItemSettingsOffAC);

            //ItemSettingsOnAC
            var newItemSettingsOnAC = getNewPowerSchemaItem(
                powerSchema,
                (s, ea) => setPowerSchemaAsOnAC(powerSchema),
                (powerSchema.Guid == configuration.Data.AutomaticPlanGuidOnAC)
                );
            MenuItem settingsOnACItem = SettingsMenuItem.Items.GetItemAt(1) as MenuItem;
            settingsOnACItem.Items.Add(newItemSettingsOnAC);
        }

        private void clearPowerSchemasInTray()
        {
            for (int i = NotifyIcon.Menu.Items.Count - 1; i >= 0; i--)
            {
                MenuItem item = NotifyIcon.Menu.Items.GetItemAt(i) as MenuItem;
                string tag = item?.Tag?.ToString();
                if (!string.IsNullOrEmpty(tag))
                    if (tag.StartsWith("pwrScheme", StringComparison.Ordinal))
                    {
                        NotifyIcon.Menu.Items.Remove(item);
                    }
            }

            MenuItem settingsOffAC = SettingsMenuItem.Items.GetItemAt(0) as MenuItem;
            MenuItem settingsOnACItem = SettingsMenuItem.Items.GetItemAt(1) as MenuItem;
            settingsOffAC.Items.Clear();
            settingsOnACItem.Items.Clear();
        }

        private MenuItem getNewPowerSchemaItem(IPowerSchema powerSchema, RoutedEventHandler clickedHandler, bool isChecked)
        {
            var newItemMain = new MenuItem()
            {
                Header = powerSchema.Name,
                Tag = $"pwrScheme{powerSchema.Guid}",
                Icon = new SymbolIcon(isChecked ? SymbolRegular.Checkmark24 : SymbolRegular.Empty),
            };
            newItemMain.Click += clickedHandler;

            return newItemMain;
        }

        # endregion

        #region AutomaticOnACSwitchRelated
        private void PwrManager_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(IPowerManager.CurrentPowerStatus)) { powerStatusChanged(); }
        }

        private void powerStatusChanged()
        {
            if (!configuration.Data.AutomaticOnACSwitch) { return; }

            var currentPowerPlugStatus = pwrManager.CurrentPowerStatus;
            Guid schemaGuidToSwitch = default(Guid);

            switch (currentPowerPlugStatus)
            {
                case PowerPlugStatus.Online:
                    schemaGuidToSwitch = configuration.Data.AutomaticPlanGuidOnAC;
                    break;
                case PowerPlugStatus.Offline:
                    schemaGuidToSwitch = configuration.Data.AutomaticPlanGuidOffAC;
                    break;
                default:
                    break;
            }

            IPowerSchema schemaToSwitchTo = pwrManager.Schemas.FirstOrDefault(sch => sch.Guid == schemaGuidToSwitch);
            if (schemaToSwitchTo == null) { return; }

            pwrManager.SetPowerSchema(schemaToSwitchTo);
        }

        #endregion

        #region OnSchemaClickMethods
        private void setPowerSchemaAsOffAC(IPowerSchema powerSchema)
        {
            configuration.Data.AutomaticPlanGuidOffAC = powerSchema.Guid;
            configuration.Save();
        }

        private void setPowerSchemaAsOnAC(IPowerSchema powerSchema)
        {
            configuration.Data.AutomaticPlanGuidOnAC = powerSchema.Guid;
            configuration.Save();
        }

        private void switchToPowerSchema(IPowerSchema powerSchema)
        {
            pwrManager.SetPowerSchema(powerSchema);
        }
        #endregion

        #region SettingsTogglesRegion
        private void EnableShortcutsToggleItem_Click(object sender, EventArgs e)
        {
            MenuItem enableShortcutsToggleItem = (MenuItem)sender;

            configuration.Data.ShowOnShortcutSwitch = !configuration.Data.ShowOnShortcutSwitch;
            enableShortcutsToggleItem.Icon = new SymbolIcon(configuration.Data.ShowOnShortcutSwitch ? SymbolRegular.Checkmark24 : SymbolRegular.Empty);
            enableShortcutsToggleItem.IsEnabled = !(Application.Current as App).HotKeyFailed;

            configuration.Save();
        }

        private void AutomaticHideItem_Click(object sender, EventArgs e)
        {
            MenuItem automaticHideItem = (MenuItem)sender;

            configuration.Data.AutomaticFlyoutHideAfterClick = !configuration.Data.AutomaticFlyoutHideAfterClick;
            automaticHideItem.Icon = new SymbolIcon(configuration.Data.AutomaticFlyoutHideAfterClick ? SymbolRegular.Checkmark24 : SymbolRegular.Empty);

            configuration.Save();
        }

        private void OnlyDefaultSchemas_Click(object sender, EventArgs e)
        {
            MenuItem onlyDefaultSchemasItem = (MenuItem)sender;

            configuration.Data.ShowOnlyDefaultSchemas = !configuration.Data.ShowOnlyDefaultSchemas;
            onlyDefaultSchemasItem.Icon = new SymbolIcon(configuration.Data.ShowOnlyDefaultSchemas ? SymbolRegular.Checkmark24 : SymbolRegular.Empty);

            configuration.Save();
        }

        private void AutomaticSwitchItem_Click(object sender, EventArgs e)
        {
            MenuItem automaticSwitchItem = (MenuItem)sender;

            configuration.Data.AutomaticOnACSwitch = !configuration.Data.AutomaticOnACSwitch;
            automaticSwitchItem.Icon = new SymbolIcon(configuration.Data.AutomaticOnACSwitch ? SymbolRegular.Checkmark24 : SymbolRegular.Empty);

            if (configuration.Data.AutomaticOnACSwitch) { powerStatusChanged(); }

            configuration.Save();
        }

        #endregion

        #region OtherItemsClicked
        private void Tray_LeftClick(Wpf.Ui.Tray.Controls.NotifyIcon sender, RoutedEventArgs e)
        {
            e.Handled = true;
            ShowFlyout?.Invoke();
        }

        private void TrayItemAbout_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(AppStrings.AboutAppURL);
        }

        private void TrayItemExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            App.Current.Shutdown();
        }
        #endregion

    }
}
