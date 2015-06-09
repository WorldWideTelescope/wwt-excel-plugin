//-----------------------------------------------------------------------
// <copyright file="WorksheetExtensionsTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using InteropExcel = Microsoft.Office.Interop.Excel;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// This is a test class for WorksheetExtensionsTest and is intended to contain all WorksheetExtensionsTest Unit Tests.
    /// </summary>
    [TestClass()]
    public class WorksheetExtensionsTest
    {
        /// <summary>
        /// Test context instance.
        /// </summary>
        private TestContext testContextInstance;

        /// <summary>
        /// Gets or sets the test context which provides information about and functionality for the current test run.
        /// </summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        /// <summary>
        /// A test for GetRange
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetRangeTest()
        {
            InteropExcel.Application application = new InteropExcel.Application();

            try
            {
                InteropExcel.Workbook book = application.OpenWorkbook("TestData.xlsx", false);

                InteropExcel._Worksheet worksheet = book.ActiveSheet as InteropExcel._Worksheet;

                // Get the named range stored in the test data excel file.
                InteropExcel.Name expected = book.Names.GetNamedRange("GetRange");

                InteropExcel.Range firstCell = worksheet.Cells[9, 5];
                int rowSize = 8;
                int columnSize = 5;
                InteropExcel.Range actual = null;

                // Get the range using the custom extension methods.
                actual = WorksheetExtensions.GetRange(worksheet, firstCell, rowSize, columnSize);

                Assert.AreEqual(expected.RefersToRange.Address, actual.Address);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for GetAffectedNamedRanges
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetAffectedNamedRangesTest()
        {
            InteropExcel.Application application = new InteropExcel.Application();

            try
            {
                InteropExcel.Workbook workbook = application.OpenWorkbook("TestData.xlsx", false);
                InteropExcel._Worksheet worksheet = workbook.ActiveSheet as InteropExcel._Worksheet;

                Dictionary<string, string> namedRanges = new Dictionary<string, string>();

                // Initialize ranges
                workbook.Names.GetNamedRange("TestRangeOne").Visible = false;
                workbook.Names.GetNamedRange("TestRangeTwo").Visible = false;
                workbook.Names.GetNamedRange("TestRangeThree").Visible = false;
                workbook.Names.GetNamedRange("TestRangeTarget").Visible = false;

                // Build the dictionary of name : address pairs
                namedRanges["TestRangeOne"] = workbook.Names.GetNamedRange("TestRangeOne").RefersToRange.Address;
                namedRanges["TestRangeTwo"] = workbook.Names.GetNamedRange("TestRangeTwo").RefersToRange.Address;
                namedRanges["TestRangeThree"] = workbook.Names.GetNamedRange("TestRangeThree").RefersToRange.Address;

                // Get the target range that will be tested for intersection with the above ranges
                InteropExcel.Range targetRange = workbook.Names.GetNamedRange("TestRangeTarget").RefersToRange;

                // Build the expected output
                Dictionary<string, string> expected = new Dictionary<string, string>();
                expected["TestRangeTwo"] = namedRanges["TestRangeTwo"];
                expected["TestRangeThree"] = namedRanges["TestRangeThree"];

                // Get the actual output
                Dictionary<string, string> actual;
                actual = worksheet.GetAffectedNamedRanges(targetRange, namedRanges);

                Assert.AreEqual(expected.Count, actual.Count);
                foreach (string rangeName in expected.Keys)
                {
                    Assert.IsTrue(actual.ContainsKey(rangeName));
                }
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for GetRangeNameForActiveCell
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetRangeNameForActiveCellTest()
        {
            InteropExcel.Application application = new InteropExcel.Application();

            try
            {
                InteropExcel.Workbook workbook = application.OpenWorkbook("TestData.xlsx", false);

                // Initialize Ranges
                workbook.Names.GetNamedRange("TestRangeOne").Visible = false;
                workbook.Names.GetNamedRange("TestRangeTwo").Visible = false;
                workbook.Names.GetNamedRange("TestRangeThree").Visible = false;
                workbook.Names.GetNamedRange("TestRangeTarget").Visible = false;

                // Get the target range that will be used to set the active sheet
                InteropExcel.Name targetName = workbook.Names.GetNamedRange("TestRangeTarget");

                // Activate the worksheet which contains the named range.
                ((_Worksheet)targetName.RefersToRange.Worksheet).Activate();
               
                InteropExcel._Worksheet worksheet = workbook.ActiveSheet as InteropExcel._Worksheet;

                // Select a cell in this sheet
                Range activeCell = targetName.RefersToRange.Cells[1, 1];
                activeCell.Select();
               
                Dictionary<string, string> namedRanges = new Dictionary<string, string>();
                
                // Build the dictionary of name : address pairs
                namedRanges["TestRangeOne"] = workbook.Names.GetNamedRange("TestRangeOne").RefersToRange.Address;
                namedRanges["TestRangeTwo"] = workbook.Names.GetNamedRange("TestRangeTwo").RefersToRange.Address;
                namedRanges["TestRangeThree"] = workbook.Names.GetNamedRange("TestRangeThree").RefersToRange.Address;
                namedRanges["TestRangeTarget"] = workbook.Names.GetNamedRange("TestRangeTarget").RefersToRange.Address;

                string expected = "TestRangeTarget";
                string actual;
                actual = WorksheetExtensions.GetRangeNameForActiveCell(worksheet, activeCell, namedRanges);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for IsSheetEmpty
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void IsSheetEmptyTest()
        {
            InteropExcel.Application application = new InteropExcel.Application();

            try
            {
                InteropExcel.Workbook workbook = application.OpenWorkbook("TestData.xlsx", false);
                
                // Get the target range that will be used to set the active sheet
                InteropExcel.Name targetName = workbook.Names.GetNamedRange("TestRangeTarget");

                // Activate the worksheet which contains the named range.
                ((_Worksheet)targetName.RefersToRange.Worksheet).Activate();

                InteropExcel._Worksheet worksheet = workbook.ActiveSheet as InteropExcel._Worksheet;

                bool expected = true;
                bool actual;
                actual = WorksheetExtensions.IsSheetEmpty(worksheet);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                application.Close();
            }
        }
    }
}