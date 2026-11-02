using System;
using System.Globalization;
using System.Windows.Data;
using Wpf.Ui.Controls;

namespace PowerSwitcher.TrayApp.Extensions
{
    public class PowerSchemaIconConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string name = string.Empty;
            if (value is IPowerSchema schema)
            {
                name = schema.Name;
            }
            else if (value is string str)
            {
                name = str;
            }

            if (!string.IsNullOrWhiteSpace(name))
            {
                var upper = name.ToUpperInvariant();
                if (upper.Contains("BALANCED")) return SymbolRegular.Gauge24;
                if (upper.Contains("HIGH") || upper.Contains("PERFORMANCE") || upper.Contains("ULTIMATE")) return SymbolRegular.Flash24;
                if (upper.Contains("SAVER")) return SymbolRegular.BatterySaver24;
            }

            return SymbolRegular.Settings24;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
