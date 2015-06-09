// -----------------------------------------------------------------------
// <copyright file="DoubleToStringConverter.cs" company="AditiTechnologies Pvt Ltd">
// Convert calss to convert the double value to string.
// </copyright>
// -----------------------------------------------------------------------

namespace Microsoft.Research.Wwt.Excel.Addin
{
    using System;
    using System.Windows.Data;

    /// <summary>
    /// Convert calss to convert the double value to string.
    /// </summary>
    public class DoubleToStringConverter : IValueConverter
    {
        /// <summary>
        /// Conver method.
        /// </summary>
        /// <param name="value">Value to convert.</param>
        /// <param name="targetType">Target type value.</param>
        /// <param name="parameter">Parameter value.</param>
        /// <param name="culture">Current culture.</param>
        /// <returns>Converted string.</returns>
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return value == null ? string.Empty : value.ToString();
        }

        /// <summary>
        /// Convert back from target to source.
        /// </summary>
        /// <param name="value">Value to convert.</param>
        /// <param name="targetType">Target type value.</param>
        /// <param name="parameter">Parameter value.</param>
        /// <param name="culture">Current culture.</param>
        /// <returns>Converted double value.</returns>
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            double result = 0.0;

            if (value != null)
            {
                if (!double.TryParse(value.ToString(), out result))
                {
                    result = 0.0;
                }
            }

            return result;
        }
    }
}
