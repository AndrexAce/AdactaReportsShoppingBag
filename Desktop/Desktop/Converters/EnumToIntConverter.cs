using Microsoft.UI.Xaml.Data;
using System;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Converters;

internal sealed partial class EnumToIntConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value == null)
            return 0;

        if (value is Enum enumValue)
        {
            return System.Convert.ToInt32(enumValue);
        }

        return 0;
    }

    public object? ConvertBack(object value, Type targetType, object parameter, string language)
    {
        if (value == null || targetType == null || !targetType.IsEnum)
            return null;

        try
        {
            return Enum.ToObject(targetType, value);
        }
        catch
        {
            return null;
        }
    }
}