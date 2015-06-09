//-----------------------------------------------------------------------
// <copyright file="WorkbookExtensions.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// This class has all the extensions classes required for Excel.Workbook class.
    /// </summary>
    public static class WorkbookExtensions
    {
        /// <summary>
        /// This function is used to created named range for the specified range.
        /// </summary>
        /// <param name="workbook">
        /// Workbook instance.
        /// </param>
        /// <param name="name">
        /// Name of the named range.
        /// </param>
        /// <param name="range">
        /// Selected range.
        /// </param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <returns>
        /// Named Range for the selected range.
        /// </returns>
        public static Name CreateNamedRange(this _Workbook workbook, string name, Range range)
        {
            Name namedRange = null;
            if (workbook != null && range != null && !string.IsNullOrEmpty(name))
            {
                try
                {
                    // Get the used range from the selected range. There may be empty rows in end of range 
                    // or EntireColumn might have been selected.
                    range = range.GetUsedRange();
                    namedRange = workbook.Names.Add(name, range, false);
                }
                catch (COMException comException)
                {
                    Logger.LogException(comException);
                    throw new CustomException(
                        Properties.Resources.CreateNamedRangeFailure,
                        comException,
                        true);
                }
            }

            return namedRange;
        }

        /// <summary>
        /// This function is used to retrieve the auto generated name for the named range which we are creating.
        /// </summary>
        /// <param name="workbook">
        /// Workbook instance.
        /// </param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <returns>
        /// Auto generated name for the selection range.
        /// </returns>
        public static string GetSelectionRangeName(this _Workbook workbook)
        {
            if (workbook != null)
            {
                string validName = GetValidName(workbook.ActiveSheet.Name) + "_";

                int latestSequence = 0;
                foreach (Name item in workbook.Names)
                {
                    if (item.Name.StartsWith(validName, StringComparison.OrdinalIgnoreCase))
                    {
                        int sequence = 0;
                        if (int.TryParse(item.Name.Substring(item.Name.LastIndexOf('_') + 1), out sequence))
                        {
                            latestSequence = (latestSequence < sequence) ? sequence : latestSequence;
                        }
                    }
                }

                return string.Format(CultureInfo.InvariantCulture, "{0}{1}", validName, ++latestSequence);
            }

            return string.Empty;
        }

        /// <summary>
        /// Extension method to add content to workbook custom xml parts
        /// </summary>
        /// <param name="workbook">
        /// workbook object
        /// </param>
        /// <param name="content">
        /// content to be added
        /// </param>
        public static void AddCustomXmlPart(this _Workbook workbook, string content, string xmlNamespace)
        {
            if (workbook != null && !string.IsNullOrEmpty(content))
            {
                // Delete existing part if present
                for (int i = 1; i <= workbook.CustomXMLParts.Count; i++)
                {
                    if (workbook.CustomXMLParts[i].XML.ToString().Contains(xmlNamespace))
                    {
                        workbook.CustomXMLParts[i].Delete();
                    }
                }

                // Add CustomXMLPart
                workbook.CustomXMLParts.Add(content, System.Type.Missing);
            }
        }

        /// <summary>
        /// Extension method to get content from workbook custom xml parts
        /// </summary>
        /// <param name="workbook">
        /// workbook object
        /// </param>
        /// <returns>
        /// Xml content from the workbook custom xml parts
        /// </returns>
        public static string GetCustomXmlPart(this _Workbook workbook, string xmlNamespace)
        {
            string content = string.Empty;
            if (workbook != null)
            {
                for (int i = 1; i <= workbook.CustomXMLParts.Count; i++)
                {
                    if (workbook.CustomXMLParts[i].XML.ToString().Contains(xmlNamespace))
                    {
                        content = workbook.CustomXMLParts[i].XML.ToString();
                        break;
                    }
                }
            }

            return content;
        }

        /// <summary>
        /// This function is used to get valid name for the range.
        /// </summary>
        /// <param name="name">
        /// Name which has to be converted.
        /// </param>
        /// <returns>
        /// Valid name representation of the input.
        /// </returns>
        private static string GetValidName(string name)
        {
            string validName = Regex.Replace(name, Constants.InvalidNameCharactersPattern, string.Empty);
            if (!string.IsNullOrEmpty(validName))
            {
                if (Regex.IsMatch(validName, Constants.StartsWithDigitOrDotPattern))
                {
                    validName = string.Format(CultureInfo.InvariantCulture, "{0}_{1}", Constants.DefaultLayerName, validName);
                }
            }
            else
            {
                validName = Constants.DefaultLayerName;
            }

            return validName;
        }
    }
}
