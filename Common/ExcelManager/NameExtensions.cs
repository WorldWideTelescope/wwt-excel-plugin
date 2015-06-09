//-----------------------------------------------------------------------
// <copyright file="NameExtensions.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// This class has all the extensions classes required for Excel.Name class.
    /// </summary>
    public static class NameExtensions
    {
        /// <summary>
        /// This function is used to check if the NamedRange is valid or not.
        /// </summary>
        /// <param name="namedRange">
        /// Name of the named range.
        /// </param>
        /// <returns>
        /// True if the named range is valid; Otherwise false.
        /// </returns>
        public static bool IsValid(this Name namedRange)
        {
            // Check if the named range is valid or not.
            bool isValid = true;
            try
            {
                if (namedRange == null || namedRange.RefersToRange == null)
                {
                    isValid = false;
                }
            }
            catch (COMException ex)
            {
                isValid = false;
                Logger.LogException(ex);
            }

            return isValid;
        }

        /// <summary>
        /// Check if it is a WWT range or not
        /// </summary>
        /// <param name="namedRange">named range</param>
        /// <returns>true or false</returns>
        public static bool IsWWTRange(this Name namedRange)
        {
            // Check if the named range is WWT range or not.
            return namedRange != null && namedRange.Visible == false;
        }

        /// <summary>
        /// Get named range from the collection
        /// </summary>
        /// <param name="namedCollection">named Collection</param>
        /// <param name="rangeName">range Name</param>
        /// <returns>Name object</returns>
        public static Name GetNamedRange(this Names namedCollection, string rangeName)
        {
            Name result = null;
            if (namedCollection != null && !string.IsNullOrEmpty(rangeName))
            {
                try
                {
                    result = namedCollection.Item(rangeName);
                }
                catch (COMException ex)
                {
                    // Ignore COMException.
                    Logger.LogException(ex);
                }
            }

            return result;
        }
    }
}
