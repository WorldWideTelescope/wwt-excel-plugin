//-----------------------------------------------------------------------
// <copyright file="TickConverter.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Converter for Tick
    /// </summary>
    public class TickConverter : IMultiValueConverter
    {
        /// <summary>
        /// Converts collection of values into Ticks
        /// </summary>
        /// <param name="values">collection of values</param>
        /// <param name="targetType">target type</param>
        /// <param name="parameter">parameter object</param>
        /// <param name="culture">current culture</param>
        /// <returns>converted object</returns>
        public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (values != null && values.Count() > 0)
            {
                int index = System.Convert.ToInt32(values[0], CultureInfo.CurrentCulture);
                Collection<double> decayValues = (Collection<double>)values[1];
                return decayValues[index - 1].ToString(CultureInfo.CurrentCulture);
            }

            return null;
        }

        /// <summary>
        /// Convert back from tick to value
        /// </summary>
        /// <param name="value">value to be returned</param>
        /// <param name="targetTypes">target types</param>
        /// <param name="parameter">parameter object</param>
        /// <param name="culture">current culture</param>
        /// <returns>converted object collection</returns>
        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
