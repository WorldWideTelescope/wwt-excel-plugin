//-----------------------------------------------------------------------
// <copyright file="WorksheetExtensions.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Microsoft.Research.Wwt.Excel.Common
{
    public static class WorksheetExtensions
    {
        /// <summary>
        /// This function is used to retrieve all the affected named ranges.
        /// </summary>
        /// <param name="worksheet">
        /// worksheet object
        /// </param>
        /// <param name="target">
        /// Target Range
        /// </param>
        /// <param name="namedRanges">
        /// Collection of ranges.
        /// </param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <returns>
        /// List of all affected ranges.
        /// </returns>
        public static Dictionary<string, string> GetAffectedNamedRanges(this _Worksheet worksheet, Range target, Dictionary<string, string> namedRanges)
        {
            Dictionary<string, string> affectedRanges = new Dictionary<string, string>();

            if (worksheet != null && target != null && namedRanges != null)
            {
                foreach (Name name in worksheet.Application.ActiveWorkbook.Names)
                {
                    try
                    {
                        // Created by WWT Add-in.
                        if (name.IsWWTRange() && namedRanges.ContainsKey(name.Name))
                        {
                            string rangeAddress = namedRanges[name.Name];
                            Range range = worksheet.Application.Range[rangeAddress];

                            // 1st Check :- 
                            //       a. If last column is deleted. 
                            //       b. Rest of scenarios like update/insert/delete rows.
                            // 2nd Check :- 
                            //       a. If last column is deleted and does an undo to revert back the change. 
                            //       b. A new column is inserted in the last and the named range has updated.
                            if (range.HasChanged(target) ||
                                name.RefersToRange.HasChanged(target) ||
                                (target.HasDirectDependants() && (range.HasChanged(target.DirectDependents) || name.RefersToRange.HasChanged(target.DirectDependents))))
                            {
                                if (!affectedRanges.ContainsKey(name.Name))
                                {
                                    affectedRanges.Add(name.Name, name.RefersTo);
                                }
                            }
                        }
                    }
                    catch (COMException ex)
                    {
                        Logger.LogException(ex);

                        // Invalid named range. 
                        // So we add it to the affectedRanges.
                        if (!affectedRanges.ContainsKey(name.Name))
                        {
                            affectedRanges.Add(name.Name, name.RefersTo);
                        }
                    }
                }
            }

            return affectedRanges;
        }

        /// <summary>
        /// Checks if the current worksheet is empty
        /// </summary>
        /// <param name="worksheet">Current worksheet</param>
        /// <returns>Is the current sheet empty</returns>
        public static bool IsSheetEmpty(this _Worksheet worksheet)
        {
            bool result = false;
            if (worksheet != null)
            {
                Range startingRow = worksheet.Cells.Find("*", SearchOrder: XlSearchOrder.xlByRows, SearchDirection: XlSearchDirection.xlPrevious);
                if (startingRow == null || startingRow.Value2 == null)
                {
                    result = true;
                }
            }
            return result;
        }

        /// <summary>
        /// Get the range of the worksheet for the defined row and column size
        /// </summary>
        /// <param name="worksheet">Current worksheet</param>
        /// <param name="firstCell">First active cell</param>
        /// <param name="rowSize">Row size of the worksheet</param>
        /// <param name="columnSize">Column size of the worksheet</param>
        /// <returns>Range for the given row and column size</returns>
        public static Range GetRange(this _Worksheet worksheet, Range firstCell, int rowSize, int columnSize)
        {
            Range range = null;
            if (worksheet != null && firstCell != null)
            {
                firstCell = firstCell.Cells[1, 1];
                if (firstCell != null)
                {
                    range = firstCell.Resize[rowSize, columnSize];
                }
            }
            return range;
        }

        /// <summary>
        /// This function is used to retrieve named range name where the active cell belongs.
        /// </summary>
        /// <param name="worksheet">
        /// worksheet object
        /// </param>
        /// <param name="activeCell">
        /// active Cell 
        /// </param>
        /// <param name="namedRanges">
        /// Collection of ranges.
        /// </param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <returns>
        /// First range name matching the criteria.
        /// </returns>
        public static string GetRangeNameForActiveCell(this _Worksheet worksheet, Range activeCell, Dictionary<string, string> namedRanges)
        {
            string rangeName = string.Empty;
            if (worksheet != null && activeCell != null && namedRanges != null)
            {
                foreach (Name name in worksheet.Application.ActiveWorkbook.Names)
                {
                    try
                    {
                        // Created by WWT Add-in.
                        if (name.IsWWTRange() && namedRanges.ContainsKey(name.Name))
                        {
                            string rangeAddress = namedRanges[name.Name];
                            Range range = worksheet.Application.Range[rangeAddress];

                            if (range.HasChanged(activeCell))
                            {
                                rangeName = name.Name;
                                break;
                            }
                        }
                    }
                    catch (COMException)
                    {
                        // Consume exception for invalid ranges
                    }
                }
            }

            return rangeName;
        }
    }
}
