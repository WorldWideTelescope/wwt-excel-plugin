//-----------------------------------------------------------------------
// <copyright file="NameExtensionsTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using Microsoft.Office.Interop.Excel;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{
    /// <summary>
    /// This is a test class for NameExtensionsTest and is intended
    /// to contain all NameExtensionsTest Unit Tests
    /// </summary>
    [TestClass()]
    public class NameExtensionsTest
    {
        /// <summary>
        /// Test context instance.
        /// </summary>
        private TestContext testContextInstance;

        /// <summary>
        /// Gets or sets the test context which provides
        /// information about and functionality for the current test run.
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
        /// A test for IsValid - Positive scenario
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void IsValidTrueTest()
        {
             Application application = new Application();

             try
             {
                 Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                 Name namedRange = book.Names.GetNamedRange("InsertRows");
                 bool expected = true;
                 bool actual = namedRange.IsValid();
                 Assert.AreEqual(expected, actual);
             }
             finally
             {
                 application.Close();
             }
        }

        /// <summary>
        /// A test for IsValid - Negative scenario
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void IsValidFalseTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                Name namedRange = book.Names.GetNamedRange("NonExistingRange");
                bool expected = false;
                bool actual = namedRange.IsValid();
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for GetNamedRange
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetNamedRangeTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                Names namedCollection = book.Names;
                Name expected = book.Names.GetNamedRange("InsertRows");

                string rangeName = "InsertRows";
                Name actual = namedCollection.GetNamedRange(rangeName);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for IsWWTRange - Positive scenario
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void IsWWTRangeTrueTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                Name namedRange = book.Names.GetNamedRange("TestProperties_1");
                bool expected = true;
                bool actual = namedRange.IsWWTRange();
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for IsWWTRange - Negative scenario
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void IsWWTRangeFalseTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                Name namedRange = book.Names.GetNamedRange("InsertRows");
                bool expected = false;
                bool actual = namedRange.IsWWTRange();
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                application.Close();
            }
        }
    }
}