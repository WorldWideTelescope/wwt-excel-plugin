//-----------------------------------------------------------------------
// <copyright file="NameExtensionsTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2010. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using InteropExcel = Microsoft.Office.Interop.Excel;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// This is a test class for NameExtensionsTest and is intended
    /// to contain all NameExtensionsTest Unit Tests
    /// </summary>
    [TestClass()]
    public class NameExtensionsTest
    {
        /// <summary>
        /// Check if an existing named range is correctly retrieved
        /// </summary>
        [TestMethod()]
        public void GetNamedRangeTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Visible = false;
                excelApp.Workbooks.Add();
                InteropExcel.Workbook workbook = excelApp.ActiveWorkbook;
                workbook.Names.Add("TestNameOne", "A1", false);
                workbook.Names.Add("TestNameTwo", "A2", false);
                Names namedCollection = workbook.Names;
                string expected = "TestNameOne";
                Name expectedName = workbook.Names.GetNamedRange("TestNameOne");
                Name actualName;
                actualName = NameExtensions.GetNamedRange(namedCollection, expected);
                Assert.AreEqual(expectedName, actualName);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Check if a WWT Range is correctly identified
        /// </summary>
        [TestMethod()]
        public void IsWWTRangeTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Visible = false;
                excelApp.Workbooks.Add();
                InteropExcel.Workbook workbook = excelApp.ActiveWorkbook;
                Name namedRange = workbook.Names.Add("TestNameOne", "A1", false);
                bool expected = true;
                bool actual = NameExtensions.IsWWTRange(namedRange);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Check if a range is correctly identified as not being a WWT Range
        /// </summary>
        [TestMethod()]
        public void IsWWTRangeTestNegative()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Visible = false;
                excelApp.Workbooks.Add();
                InteropExcel.Workbook workbook = excelApp.ActiveWorkbook;
                Name namedRange = workbook.Names.Add("TestNameOne", "A1", true);
                bool expected = false;
                bool actual = NameExtensions.IsWWTRange(namedRange);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }
    }
}
