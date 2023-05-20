using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace WordRename
{
    [ValueConversion(typeof(Boolean), typeof(Visibility))]
    internal class EnableToVisibleConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if ((bool)value) { return Visibility.Hidden; }
            else { return Visibility.Visible; }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var d = (Visibility)value;
            if (d == Visibility.Hidden) { return false; } else { return true; }
        }
    }
}
