//-----------------------------------------------------------------------
// <copyright file="RangeExtensions.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// This class has all the extensions classes required for Excel.Range class.
    /// </summary>
    public static class RangeExtensions
    {
        /// <summary>
        /// This function is used to check if the selected areas in the range belong to same set of columns.
        /// </summary>
        /// <param name="range">
        /// Selected range.
        /// </param>
        /// <returns>
        /// True if the areas belongs to the same column set; Otherwise false.
        /// </returns>
        public static bool IsValid(this Range range)
        {
            bool isValidRanges = false;
            if (range != null)
            {
                isValidRanges = true;
                Areas areas = range.Areas; // Returns all the areas selected in the range.
                if (areas.Count > 1)
                {
                    for (int index = 1; index < areas.Count; index++)
                    {
                        // Logic:
                        //      Get the 1st and 2nd Areas in the list of areas.
                        //      Then get the address for entire column ranges of the 1st and 2nd areas.
                        //      If both address are equal then the areas belong to the same set of columns.
                        Range areaRange1 = areas[index];
                        Range areaRange2 = areas[index + 1];
                        string address1 = areaRange1.EntireColumn.Address;
                        string address2 = areaRange2.EntireColumn.Address;
                        string address1NextEntireColumn = areaRange1.Next.EntireColumn.Address;

                        // If both the areas does not belong to same set of columns, then both the areas must be entire column and 
                        // also in sequence, next to each other.
                        // ContainsCompleteColumn() will make sure that both area1 and area2 are EntireColumn selection and area2 must
                        // be next column to area1. Entire columns must be in sequence.
                        if (address1 != address2 && (!areaRange1.ContainsCompleteColumn() || !areaRange2.ContainsCompleteColumn() || address2 != address1NextEntireColumn))
                        {
                            isValidRanges = false;
                            break;
                        }
                    }
                }
            }

            return isValidRanges ? !range.ContainsCompleteRow() : isValidRanges;
        }

        /// <summary>
        /// Check whether the range contains a complete row in any of its areas
        /// </summary>
        /// <param name="range">The ranges instance on which this method is called</param>
        /// <returns>True if the range contains an area with a complete row in it</returns>
        public static bool ContainsCompleteRow(this Range range)
        {
            bool containsEntireRow = false;
            string regexString = Constants.EntireRowPattern;
            Regex addressRegex = new Regex(regexString);
            if (range != null)
            {
                foreach (Range area in range.Areas)
                {
                    if (addressRegex.IsMatch(area.Address))
                    {
                        containsEntireRow = true;
                        break;
                    }
                }
            }

            return containsEntireRow;
        }

        /// <summary>
        /// Check whether the range contains a complete column in any of its areas
        /// </summary>
        /// <param name="range">The ranges instance on which this method is called</param>
        /// <returns>True if the range contains an area with a complete column in it</returns>
        public static bool ContainsCompleteColumn(this Range range)
        {
            bool containsEntireColumn = false;
            string regexString = Constants.EntireColumnPattern;
            Regex addressRegex = new Regex(regexString);
            if (range != null)
            {
                foreach (Range area in range.Areas)
                {
                    if (addressRegex.IsMatch(area.Address))
                    {
                        containsEntireColumn = true;
                        break;
                    }
                }
            }

            return containsEntireColumn;
        }

        /// <summary>
        /// Gets the used range from the given range using the intersection with the UsedRange property of the
        /// parent worksheet to which the range belongs to.
        /// </summary>
        /// <param name="range">Range object for which used range to be obtained.</param>
        /// <returns>Used range of the range object.</returns>
        public static Range GetUsedRange(this Range range)
        {
            Range usedRange = null;

            if (range != null)
            {
                if (range.ContainsCompleteColumn())
                {
                    usedRange = range.Worksheet.Application.Intersect(range, range.Worksheet.UsedRange);

                    // If entire worksheet is empty, usedRange will be null.
                    if (usedRange != null)
                    {
                        int emptyRowCount = 0;

                        for (int i = usedRange.Rows.Count; i >= 1; i--)
                        {
                            Range rowRange = usedRange.Rows[i] as Range;
                            Range foundRange = null;

                            try
                            {
                                foundRange = rowRange.SpecialCells(XlCellType.xlCellTypeBlanks);
                            }
                            catch (COMException)
                            {
                                // If there are no cell with valid value in the range, then COMExcpetion is expected.
                                break;
                            }

                            // Got the last non-empty row.
                            if (foundRange.Address != rowRange.Address)
                            {
                                break;
                            }

                            emptyRowCount++;
                        }

                        usedRange = usedRange.Resize[usedRange.Rows.Count - emptyRowCount, Type.Missing];
                        usedRange.Select();
                    }
                }

                // In case if worksheet is empty or there are no entire columns, set the input range as usedRange.
                if (usedRange == null)
                {
                    usedRange = range;
                }
            }

            return usedRange;
        }

        /// <summary>
        /// This function is used to get the list of headers specified in the current selection range.
        /// </summary>
        /// <param name="range">
        /// Excel range which is in scope.
        /// </param>
        /// <returns>
        /// Header data as List.
        /// </returns>
        public static Collection<string> GetHeader(this Range range)
        {
            var headerValues = new Collection<string>();
            if (range != null && range.IsValid())
            {
                object[,] data = range.GetDataArray(true);
                if (data != null)
                {
                    int cols = data.GetLength(1);
                    int row = 1; // Considering first row is always the header.

                    for (int col = 1; col <= cols; col++)
                    {
                        object value = data[row, col];
                        headerValues.Add(value != null ? value.ToString() : string.Empty);
                    }
                }
            }

            return headerValues;
        }

        /// <summary>
        /// This function is used to get the first row data from the range.
        /// </summary>
        /// <param name="range">
        /// Excel range which is in scope.
        /// </param>
        /// <returns>First Row Data.
        /// </returns>
        public static Collection<string> GetFirstDataRow(this Range range)
        {
            var values = new Collection<string>();
            if (range != null && range.IsValid())
            {
                object[,] data = null;
                int row = 0;

                if (range.Areas[1].Rows.Count > 1)
                {
                    Range firstTwoRows = range.Areas[1].Resize[2, Type.Missing];
                    data = firstTwoRows.GetDataArray(false);
                    row = 2; // Considering first row is always the header.
                }
                else if (range.Areas.Count > 1)
                {
                    Range firstRow = range.Areas[2].Resize[1, Type.Missing];
                    data = firstRow.GetDataArray(false);
                    row = 1; // first row is always data.
                }
                else
                {
                    return null;
                }

                if (data != null)
                {
                    int colCount = data.GetLength(1);
                    for (int col = 1; col <= colCount; col++)
                    {
                        object value = data[row, col];
                        values.Add(value != null ? value.ToString() : string.Empty);
                    }
                }
            }

            return values;
        }

        /// <summary>
        /// This function can be used to retrieve the data for the selected range in array format.
        /// </summary>
        /// <param name="range">
        /// Excel range which is in scope.
        /// </param>
        /// <param name="getHeaderOnly">
        /// True to return only the first row in the range(header), false to return the entire range
        /// </param>
        /// <returns>
        /// Data in two dimensional array format.
        /// </returns>
        [SuppressMessage("Microsoft.Performance", "CA1814:PreferJaggedArraysOverMultidimensional", MessageId = "Return", Justification = "We cannot use jagged array in this scenario because the excel Object model is designed to convert the value as [,].")]
        public static object[,] GetDataArray(this Range range, bool getHeaderOnly)
        {
            object[,] obj = null;
            if (range != null && range.IsValid())
            {
                if (range.Cells.Count == 1)
                {
                    // Initialize an object array with the lower bound from 1.
                    obj = (object[,])Array.CreateInstance(typeof(object), new int[] { 1, 1 }, new int[] { 1, 1 });
                    obj[1, 1] = range.Value;
                }
                else
                {
                    if (getHeaderOnly)
                    {
                        obj = (object[,])Array.CreateInstance(typeof(object), new int[] { 1, range.EntireColumn.Count }, new int[] { 1, 1 });
                        Range headerRow = range.Rows[1];
                        if (range.EntireColumn.Count == 1)
                        {
                            obj[1, 1] = headerRow.Value;
                        }
                        else
                        {
                            obj = headerRow.Value;
                        }
                    }
                    else
                    {
                        obj = range.Value;
                    }
                }
            }

            return obj;
        }

        /// <summary>
        /// This function is used to check if the range in focus has been affected by the target range.
        /// </summary>
        /// <param name="range">
        /// Range in focus.
        /// </param>
        /// <param name="target">
        /// Target range.
        /// </param>
        /// <returns>
        /// True if the Range has changed; Otherwise false.</returns>
        public static bool HasChanged(this Range range, Range target)
        {
            bool hasChanged = false;
            if (range != null && range.IsValid() && target != null && range.Worksheet.Name == target.Worksheet.Name)
            {
                if (range.Areas.Count > 1)
                {
                    for (int index = 1; index <= range.Areas.Count; index++)
                    {
                        Range areaRange = range.Areas[index];
                        Range ranSelect = range.Worksheet.Application.Intersect(areaRange, target);
                        if (ranSelect != null)
                        {
                            hasChanged = true;
                            break;
                        }
                    }
                }
                else
                {
                    Range ranSelect = range.Worksheet.Application.Intersect(range, target);
                    if (ranSelect != null)
                    {
                        hasChanged = true;
                    }
                }
            }

            return hasChanged;
        }

        /// <summary>
        /// This function is used to retrieve the data for WWT from the range specified.
        /// </summary>
        /// <param name="range">
        /// Excel range which is in scope.
        /// </param>
        /// <returns>
        /// Data for WWT in format of string array.
        /// </returns>
        public static string[] GetData(this Range range)
        {
            List<string> dataForRanges = new List<string>();
            if (range != null && range.IsValid())
            {
                Areas areas = range.Areas;
                if (areas.Count > 1)
                {
                    for (int index = 1; index <= areas.Count; index++)
                    {
                        string areaData = GetDataFromRange(areas[index]);
                        if (!string.IsNullOrEmpty(areaData))
                        {
                            dataForRanges.Add(areaData);
                        }
                    }
                }
                else
                {
                    dataForRanges.Add(GetDataFromRange(range));
                }
            }

            return dataForRanges.ToArray();
        }

        /// <summary>
        /// This function is used to check if there are any direct dependants of the range.
        /// </summary>
        /// <param name="range">
        /// Excel range which is in scope.
        /// </param>
        /// <returns>
        /// True if the Range has Direct Dependants; Otherwise false.
        /// </returns>
        public static bool HasDirectDependants(this Range range)
        {
            bool hasDirectDenpendants = false;

            try
            {
                hasDirectDenpendants = range != null && range.DirectDependents.Count > 0;
            }
            catch (COMException ex)
            {
                // No Direct dependents
                hasDirectDenpendants = false;
                Logger.LogException(ex);
            }

            return hasDirectDenpendants;
        }

        /// <summary>
        /// Validates if range has formula in any of the cells
        /// </summary>
        /// <param name="currentRange">Current selected range</param>
        /// <returns>True if any of all cells of range has formula</returns>
        public static bool ValidateFormula(this Range currentRange)
        {
            bool hasFormula = true;
            if (currentRange != null)
            {
                if (!string.IsNullOrEmpty(currentRange.HasFormula.ToString()))
                {
                    bool.TryParse(currentRange.HasFormula.ToString(), out hasFormula);
                }
            }

            return hasFormula;
        }

        /// <summary>
        /// Gets the rows count for the given range. Range can have more than one Area, and rows count should be getting
        /// count of rows in all the areas.
        /// </summary>
        /// <param name="currentRange">Range object</param>
        /// <returns>Number of rows</returns>
        public static int GetRowsCount(this Range currentRange)
        {
            int rowsCount = 0;

            if (currentRange != null)
            {
                foreach (Range area in currentRange.Areas)
                {
                    rowsCount += area.Rows.Count;
                }
            }

            return rowsCount;
        }

        /// <summary>
        /// Sets the given value in the given Range. There could be more than one Area in an Range. Value should be set accordingly.
        /// </summary>
        /// <param name="currentRange">Current range to which value has to be set.</param>
        /// <param name="value">Value to be set.</param>
        [SuppressMessage("Microsoft.Performance", "CA1814:PreferJaggedArraysOverMultidimensional", Justification = "The excel Object model is designed to convert the value as [,].")]
        public static void SetValue(this Range currentRange, object[,] value)
        {
            if (currentRange != null)
            {
                if (value != null)
                {
                    int rowIndex = 0;
                    int columnsCount = value.GetLength(1);
                    foreach (Range area in currentRange.Areas)
                    {
                        object[,] tempValue = new object[area.Rows.Count, columnsCount];
                        Array.Copy(value, rowIndex, tempValue, 0, area.Rows.Count * columnsCount);
                        area.Value2 = tempValue;
                        rowIndex += area.Rows.Count * columnsCount;
                    }
                }
                else
                {
                    // In case if value is null, clear any existing values.
                    currentRange.Cells.Clear();
                }
            }
        }

        /// <summary>
        /// This function is used to extract the data from the range object.
        /// </summary>
        /// <param name="range">
        /// Excel range which is in scope.
        /// </param>
        /// <returns>
        /// Extracted Data.
        /// </returns>
        private static string GetDataFromRange(Range range)
        {
            object[,] data = range.GetDataArray(false);

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

            // We need to push all the data to WWT including the header.
            int rowStart = 1;

            StringBuilder sb = new StringBuilder();
            for (int row = rowStart; row <= rows; row++)
            {
                for (int col = 1; col <= cols; col++)
                {
                    object cell = data[row, col];
                    if (col != 1)
                    {
                        sb.Append("\t");
                    }

                    if (cell != null)
                    {
                        sb.Append(cell.ToString());
                    }
                    else
                    {
                        sb.Append(" ");
                    }
                }

                sb.AppendLine(string.Empty);
            }

            return sb.ToString();
        }
    }
}
