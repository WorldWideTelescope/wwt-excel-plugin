//-----------------------------------------------------------------------
// <copyright file="ApplicationExtensions.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.IO;
using InteropExcel = Microsoft.Office.Interop.Excel;

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Class having extension methods for Excel application
    /// </summary>
    public static class ApplicationExtensions
    {
        /// <summary>
        /// Opens the given Excel workbook whose name is passed and sets the application and book objects.
        /// </summary>
        /// <param name="application">Excel application object</param>
        /// <param name="fileName">File to be opened</param>
        /// <param name="visibility">Make the excel application visible or not</param>
        /// <returns>Excel workbook object</returns>
        public static InteropExcel.Workbook OpenWorkbook(this InteropExcel.Application application, string fileName, bool visibility)
        {
            InteropExcel.Workbook workBook = null;

            if (application != null)
            {
                application.Visible = visibility;

                // Open the unit test data excel file.
                string filePath = Path.Combine(Environment.CurrentDirectory, fileName);
                workBook = application.Workbooks.Open(filePath);
            }

            return workBook;
        }

        /// <summary>
        /// Closes the workbook and the excel application.
        /// </summary>
        /// <param name="application">Excel application object</param>
        public static void Close(this InteropExcel._Application application)
        {
            if (application != null)
            {
                foreach (InteropExcel.Workbook workbook in application.Workbooks)
                {
                    workbook.Close(false);
                }

                application.Quit();
            }
        }
    }
}
