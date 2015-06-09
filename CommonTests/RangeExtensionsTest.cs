//-----------------------------------------------------------------------
// <copyright file="RangeExtensionsTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2010. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using System;
using System.Collections.ObjectModel;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using InteropExcel = Microsoft.Office.Interop.Excel;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// This is a test class for RangeExtensionsTest and is intended
    /// to contain all RangeExtensionsTest Unit Tests
    /// </summary>
    [TestClass()]
    public class RangeExtensionsTest
    {
        /// <summary>
        /// Check if data array is correctly retrieved when there is only one cell in the range
        /// </summary>
        [TestMethod()]
        public void GetDataArraySingleCellRangeTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", Type.Missing);
                range.Value2 = "TestValue";
                object[,] expected = (object[,])Array.CreateInstance(typeof(object), new int[] { 1, 1 }, new int[] { 1, 1 });
                expected[1, 1] = "TestValue";
                object[,] actual = range.GetDataArray(false);

                Assert.AreEqual(expected.Length, actual.Length);
                Assert.AreEqual(expected[1, 1], actual[1, 1]);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Check if data array is correctly retrieved when there are two cells in the range
        /// </summary>
        [TestMethod()]
        public void GetDataArrayMultipleCellRangeTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", "B1");
                Range cellOne = range.Cells[1, 1];
                cellOne.Value = "TestValue1";
                Range cellTwo = range.Cells[1, 2];
                cellTwo.Value = "TestValue2";

                object[,] expected = (object[,])Array.CreateInstance(typeof(object), new int[] { 1, 2 }, new int[] { 1, 1 });
                expected[1, 1] = "TestValue1";
                expected[1, 2] = "TestValue2";
                object[,] actual = range.GetDataArray(false);

                Assert.AreEqual(expected.Length, actual.Length);
                Assert.AreEqual(expected[1, 1], actual[1, 1]);
                Assert.AreEqual(expected[1, 2], actual[1, 2]);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Check for null range value. Exception is thrown by default
        /// </summary>
        [TestMethod()]
        public void GetDataArrayNullRangeTest()
        {
            Range range = null;
            object[,] expected = null;
            object[,] actual = RangeExtensions.GetDataArray(range, false);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// Verify that function returns null when the range is not valid
        /// </summary>
        [TestMethod()]
        public void GetDataArrayInvalidRangeTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", "I7");
                Range test = worksheet.get_Range("B10", "J15");
                Range mergedRange = excelApp.Union(range, test);

                object[,] expected = null;
                object[,] actual = mergedRange.GetDataArray(false);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Check if header values are null when there is no data in the range
        /// </summary>
        [TestMethod()]
        public void GetHeaderNullDataTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", "B1");
                range.Value2 = null;
                Collection<string> expected = new Collection<string>() { string.Empty, string.Empty };
                Collection<string> actual = range.GetHeader();
                Assert.AreEqual(actual.Count, 2);
                Assert.AreEqual(expected[0], actual[0]);
                Assert.AreEqual(expected[1], actual[1]);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Function returns an empty collection when the range is invalid
        /// </summary>
        [TestMethod()]
        public void GetHeaderInvalidRangeTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", "I7");
                Range anotherRange = worksheet.get_Range("B10", "J15");
                Range mergedRange = excelApp.Union(range, anotherRange);
                Collection<string> actual = mergedRange.GetHeader();
                Assert.AreEqual(0, actual.Count);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Check whether the function returns the expected value when there is only one data value in the range
        /// </summary>
        [TestMethod()]
        public void GetHeaderSingleCellDataTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", Type.Missing);
                range.Value2 = "TestValue";

                Collection<string> expected = new Collection<string>() { "TestValue" };
                Collection<string> actual = range.GetHeader();

                Assert.AreEqual(expected.Count, actual.Count);
                Assert.AreEqual(expected[0], actual[0]);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Check if a range containing a single cell is identified as a valid range
        /// </summary>
        [TestMethod()]
        public void IsValidSingleAreaTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", Type.Missing);
                bool expected = true;
                bool actual = range.IsValid();
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Check if a range containing a single cell is identified as a valid range
        /// </summary>
        [TestMethod()]
        public void IsValidTwoAreasPositiveTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", "I7");
                bool expected = true;

                Range test = worksheet.get_Range("A10", "I15");
                Range mergedRange = excelApp.Union(range, test);
                bool actual = mergedRange.IsValid();
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Check if a range containing a single cell is identified as a valid range
        /// </summary>
        [TestMethod()]
        public void IsValidTwoAreasNegativeTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", "I7");
                bool expected = false;

                Range test = worksheet.get_Range("B10", "J15");
                Range mergedRange = excelApp.Union(range, test);
                bool actual = mergedRange.IsValid();

                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Function should return false when the target range is null
        /// </summary>
        [TestMethod()]
        public void HasChangedNullArgumentsTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", Type.Missing);
                Range target = null;
                bool expected = false;
                bool actual = range.HasChanged(target);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Function should return false when the target range is invalid
        /// </summary>
        [TestMethod()]
        public void HasChangedInvalidRangeTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", "I7");
                Range test = worksheet.get_Range("B10", "J15");
                Range mergedRange = excelApp.Union(range, test);

                Range target = worksheet.get_Range("A1", Type.Missing);
                bool expected = false;
                bool actual = mergedRange.HasChanged(target);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Function should return false when the target range and
        /// the calling range are on different worksheets
        /// </summary>
        [TestMethod()]
        public void HasChangedRangesOnDifferentWorksheetsTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                InteropExcel.Workbook workbook = excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", "I7");
                Worksheet secondSheet = workbook.Worksheets[2];
                Range target = secondSheet.get_Range("C1", "E5");
                bool expected = false;
                bool actual = range.HasChanged(target);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Function should return true when the target range overlaps with the calling range
        /// </summary>
        [TestMethod()]
        public void HasChangedIntersectionPositiveSingleAreaTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", "I7");
                Range target = worksheet.get_Range("C1", "E5");
                bool expected = true;
                bool actual = range.HasChanged(target);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Function should return false when the target range overlaps with the calling range
        /// </summary>
        [TestMethod()]
        public void HasChangedIntersectionNegativeSingleAreaTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", "I7");
                Range target = worksheet.get_Range("A9", "I15");
                bool expected = false;
                bool actual = range.HasChanged(target);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Function should return true when the target range overlaps with the calling range
        /// </summary>
        [TestMethod()]
        public void HasChangedIntersectionPositiveMultipleAreasTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range areaOne = worksheet.get_Range("A1", "I7");
                Range areaTwo = worksheet.get_Range("A10", "I15");
                Range mergedRange = excelApp.Union(areaOne, areaTwo);

                Range target = worksheet.get_Range("A11", "I13");
                bool expected = true;
                bool actual = mergedRange.HasChanged(target);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Function should return false when the target range does not overlap with the calling range
        /// </summary>
        [TestMethod()]
        public void HasChangedIntersectionNegativeMultipleAreasTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                // worksheet names
                Range areaOne = worksheet.get_Range("A1", "I7");
                Range areaTwo = worksheet.get_Range("A10", "I15");
                Range mergedRange = excelApp.Union(areaOne, areaTwo);

                Range target = worksheet.get_Range("A8", "I8");
                bool expected = false;
                bool actual = mergedRange.HasChanged(target);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Function should return true when the target range overlaps with the calling range
        /// Here, the calling range is a union of single and multiple areas
        /// </summary>
        [TestMethod()]
        public void HasChangedIntersectionPositiveSingleAndMultipleAreasTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range areaOne = worksheet.get_Range("A1", "I7");
                Range areaTwo = worksheet.get_Range("A10", "I15");
                Range mergedRange = excelApp.Union(areaOne, areaTwo);
                Range areaThree = worksheet.get_Range("D12", Type.Missing);
                Range mergedWithSingle = excelApp.Union(mergedRange, areaThree);

                Range target = worksheet.get_Range("A11", "I13");
                bool expected = true;
                bool actual = mergedWithSingle.HasChanged(target);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Function should return true when the target range overlaps with the calling range
        /// Here, the calling range is a union of single and multiple areas
        /// </summary>
        [TestMethod()]
        public void HasChangedIntersectionPositiveSingleAndMultipleAreasIntersectSingleTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range areaOne = worksheet.get_Range("A1", "I7");
                Range areaTwo = worksheet.get_Range("A10", "I15");
                Range mergedRange = excelApp.Union(areaOne, areaTwo);
                Range areaThree = worksheet.get_Range("D12", Type.Missing);
                Range mergedWithSingle = excelApp.Union(mergedRange, areaThree);

                Range target = worksheet.get_Range("D12", Type.Missing);
                bool expected = true;
                bool actual = mergedWithSingle.HasChanged(target);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Function should return false when the target range does not overlap with the calling range
        /// Here, the calling range is a union of single and multiple areas
        /// </summary>
        [TestMethod()]
        public void HasChangedIntersectionNegativeSingleAndMultipleAreasTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range areaOne = worksheet.get_Range("A1", "I7");
                Range areaTwo = worksheet.get_Range("A10", "I15");
                Range mergedRange = excelApp.Union(areaOne, areaTwo);
                Range areaThree = worksheet.get_Range("D12", Type.Missing);
                Range mergedWithSingle = excelApp.Union(mergedRange, areaThree);

                Range target = worksheet.get_Range("A8", "I8");
                bool expected = false;
                bool actual = mergedWithSingle.HasChanged(target);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Verify that the correct data value is retrieved when there are only two cells in the range
        /// </summary>
        [TestMethod()]
        public void GetDataTwoCellsTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", "A2");
                Range cell = range.Cells[1, 1];
                cell.Value = "TestValueOne";
                cell = range.Cells[2, 1];
                cell.Value = "TestValueTwo";
                string[] expected = new string[] { "TestValueOne" + Environment.NewLine + "TestValueTwo" + Environment.NewLine };
                string[] actual = range.GetData();
                Assert.AreEqual(actual.Length, 1);
                Assert.AreEqual(expected[0], actual[0]);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Verify that the function gets data correctly when there are three rows
        /// </summary>
        [TestMethod()]
        public void GetDataFromRangePositiveTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;

                Range range = worksheet.get_Range("A1", "B3");
                Range cell = range.Cells[1, 1];
                cell.Value = "HeaderValueOne";
                cell = range.Cells[1, 2];
                cell.Value = "HeaderValueTwo";
                cell = range.Cells[2, 1];
                cell.Value = "DataValueOne";
                cell = range.Cells[2, 2];
                cell.Value = "DataValueTwo";
                cell = range.Cells[3, 1];
                cell.Value = "DataValueThree";
                cell = range.Cells[3, 2];
                cell.Value = "DataValueFour";

                string expected = "HeaderValueOne\tHeaderValueTwo" + Environment.NewLine + "DataValueOne\tDataValueTwo" + Environment.NewLine + "DataValueThree\tDataValueFour" + Environment.NewLine;
                string actual = RangeExtensions_Accessor.GetDataFromRange(range);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }
        
        /// <summary>
        /// A test for ValidateFormula
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void ValidateFormulaTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);

                // Get the named range stored in the test data excel file.
                // This range contains no formula.
                Name rangeWithoutFormula = book.Names.GetNamedRange("RangeWithoutFormula");

                bool expected = false;
                bool actual = RangeExtensions.ValidateFormula(rangeWithoutFormula.RefersToRange);
                Assert.AreEqual(expected, actual);

                // Get the named range stored in the test data excel file.
                // This range contains formula.
                Name rangeWithFormula = book.Names.GetNamedRange("RangeWithFormula");

                expected = true;
                actual = RangeExtensions.ValidateFormula(rangeWithFormula.RefersToRange);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for GetFirstDataRow
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetFirstDataRowTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                InteropExcel.Workbook workbook = excelApp.OpenWorkbook("TestData.xlsx", false);

                // Get the target range that will be used to set the active sheet
                InteropExcel.Name targetName = workbook.Names.GetNamedRange("DataRangeOne");

                Range range = targetName.RefersToRange;
                Collection<string> expected = new Collection<string>() { "23", "34", "3", "3/7/2011 12:00:00 AM" };
                Collection<string> actual;
                actual = range.GetFirstDataRow();
                Assert.AreEqual(expected.Count, actual.Count);

                for (int i = 0; i < expected.Count; i++)
                {
                    Assert.AreEqual(expected[i], actual[i]);
                }
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// A test for GetFirstDataRow
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetFirstDataRowTwoAreasTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                InteropExcel.Workbook workbook = excelApp.OpenWorkbook("TestData.xlsx", false);

                // Get the target range that will be used to set the active sheet
                InteropExcel.Name targetName = workbook.Names.GetNamedRange("DataRangeTwo");

                Range range = targetName.RefersToRange;
                Collection<string> expected = new Collection<string>() { "45", "63", "4.6", "3/7/2011 1:39:45 PM" };
                Collection<string> actual;
                actual = range.GetFirstDataRow();
                Assert.AreEqual(expected.Count, actual.Count);
                
                for (int i = 0; i < expected.Count; i++)
                {
                    Assert.AreEqual(expected[i], actual[i]);
                }
            }
            finally
            {
                excelApp.Close();
            }
        }
        
        /// <summary>
        /// A test for GetData
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetData()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                InteropExcel.Workbook workbook = excelApp.OpenWorkbook("TestData.xlsx", false);

                // Get the target range 
                InteropExcel.Name targetName = workbook.Names.GetNamedRange("DataRangeTwo");

                string[] expected = new string[] { "23\t34\t3\t3/7/2011 12:00:00 AM\r\n", "45\t63\t4.6\t3/7/2011 1:39:45 PM\r\n67\t32\t5.3\t2/1/2009 12:00:00 AM\r\n" };
                string[] actual = targetName.RefersToRange.GetData();

                Assert.AreEqual(expected.Length, actual.Length);
                Assert.AreEqual(expected[0], actual[0]);
                Assert.AreEqual(expected[1], actual[1]);
            }
            finally
            {
                excelApp.Close();
            }
        }
    }
}
