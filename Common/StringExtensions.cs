//-----------------------------------------------------------------------
// <copyright file="StringExtensions.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics.CodeAnalysis;

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// This class has extensions methods for string class.
    /// </summary>
    public static class StringExtensions
    {
        /// <summary>
        /// Converts the specified string to a integer using TryParse.
        /// </summary>
        /// <param name="value">
        /// The string to convert.
        /// </param>
        /// <param name="defaultValue">
        /// The default value for if the value cannot be parsed.
        /// </param>
        /// <returns>
        /// The specified string as a int.
        /// </returns>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "integer", Justification = "We are converting from string to integer, so we need to have name identifies the same.")]
        public static int AsInteger(this string value, int defaultValue)
        {
            int temp;
            temp = int.TryParse(value, out temp) ? temp : defaultValue;
            return temp;
        }

        /// <summary>
        /// Converts the specified string to a double using TryParse.
        /// </summary>
        /// <param name="value">
        /// The string to convert.
        /// </param>
        /// <param name="defaultValue">
        /// The default value for if the value cannot be parsed.
        /// </param>
        /// <returns>
        /// The specified string as a double.
        /// </returns>
        public static double AsDouble(this string value, double defaultValue)
        {
            double temp;
            temp = double.TryParse(value, out temp) ? temp : defaultValue;
            return temp;
        }

        /// <summary>
        /// Converts the specified string to a enumeration using TryParse.
        /// </summary>
        /// <param name="value">
        /// The string to convert.
        /// </param>
        /// <param name="defaultValue">
        /// The default value for if the value cannot be parsed.
        /// </param>
        /// <returns>
        /// The specified string as a enumeration.
        /// </returns>
        public static T AsEnum<T>(this string value, T defaultValue) where T : struct
        {
            T temp = defaultValue;
            temp = Enum.TryParse<T>(value, out temp) ? temp : defaultValue;
            return temp;
        }

        /// <summary>
        /// Converts the specified string to a DateTime using TryParse.
        /// </summary>
        /// <param name="value">
        /// The string to convert.
        /// </param>
        /// <param name="defaultValue">
        /// The default value for if the value cannot be parsed.
        /// </param>
        /// <returns>
        /// The specified string as a DateTime.
        /// </returns>
        public static DateTime AsDateTime(this string value, DateTime defaultValue)
        {
            DateTime temp;
            temp = DateTime.TryParse(value, out temp) ? temp : defaultValue;
            return temp;
        }

        /// <summary>
        /// Converts the specified string to a Boolean using TryParse.
        /// </summary>
        /// <param name="value">
        /// The string to convert.
        /// </param>
        /// <param name="defaultValue">
        /// The default value for if the value cannot be parsed.
        /// </param>
        /// <returns>
        /// The specified string as a Boolean.
        /// </returns>
        public static bool AsBoolean(this string value, bool defaultValue)
        {
            bool temp;
            temp = bool.TryParse(value, out temp) ? temp : defaultValue;
            return temp;
        }

        /// <summary>
        /// Converts the specified string to a TimeSpan using TryParse.
        /// </summary>
        /// <param name="value">
        /// The string to convert.
        /// </param>
        /// <param name="defaultValue">
        /// The default value for if the value cannot be parsed.
        /// </param>
        /// <returns>
        /// The specified string as a TimeSpan.
        /// </returns>
        public static TimeSpan AsTimeSpan(this string value, TimeSpan defaultValue)
        {
            TimeSpan temp;
            temp = TimeSpan.TryParse(value, out temp) ? temp : defaultValue;
            return temp;
        }
    }
}
