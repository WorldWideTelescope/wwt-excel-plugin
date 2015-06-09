//-----------------------------------------------------------------------
// <copyright file="WorkbookExtensionsTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2010. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using System;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using InteropExcel = Microsoft.Office.Interop.Excel;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// This is a test class for WorkbookExtensionsTest and is intended
    /// to contain all WorkbookExtensionsTest Unit Tests
    /// </summary>
    [TestClass()]
    public class WorkbookExtensionsTest
    {  
        /// <summary>
        /// Verify that an empty xml string added as a custom xml part does not create a custom xml part
        /// </summary>
        [TestMethod()]
        public void AddCustomXmlPartTestEmptyString()
        {   
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Workbook workbook = excelApp.Workbooks.Add();

                string content = string.Empty;
                int expected = workbook.CustomXMLParts.Count;
                workbook.AddCustomXmlPart(content, Common.Constants.XmlNamespace);
                Assert.AreEqual(workbook.CustomXMLParts.Count, expected);
                string existingContent = workbook.GetCustomXmlPart(Common.Constants.XmlNamespace);
                Assert.AreEqual(content, existingContent);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Verify that a custom xml part can be added and retrieved correctly
        /// </summary>
        [TestMethod()]
        public void AddCustomXmlPartTestOnePart()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Workbook workbook = excelApp.Workbooks.Add();

                string content = "<Product xmlns:ns=\"" + Common.Constants.XmlNamespace + "\">" + "<ProductName>WWT Excel Add-In</ProductName>" + "</Product>";
                workbook.AddCustomXmlPart(content, Common.Constants.XmlNamespace);
                string existingContent = workbook.GetCustomXmlPart(Common.Constants.XmlNamespace);
                Assert.AreEqual(content, existingContent);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Verify that an invalid xml string cannot be added to the workbook
        /// </summary>
        [TestMethod()]
        [ExpectedException(typeof(System.Runtime.InteropServices.COMException))]
        public void AddCustomXmlPartTestOnePartInvalidXML()
        {
            string content = "<Product xmlns:ns=\"" + Common.Constants.XmlNamespace + "\"></Product>" + "<ProductName>WWT Excel Add-In</ProductName>";
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                excelApp.Workbooks.Add();
                InteropExcel.Workbook workbook = excelApp.Workbooks.Add();
                workbook.AddCustomXmlPart(content, Common.Constants.XmlNamespace);
                Assert.Fail("Invalid XML is inserted in to custom XML part!");
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Verify that only the latest custom xml part gets maintained in the workbook
        /// </summary>
        [TestMethod()]
        public void AddCustomXmlPartTestTwoParts()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {   
                InteropExcel.Workbook workbook = excelApp.Workbooks.Add();
                string content = "<Product xmlns:ns=\"" + Common.Constants.XmlNamespace + "\">" + "<ProductName>WWT Excel Add-In</ProductName>" + "</Product>";
                string moreContent = "<Product xmlns:ns=\"" + Common.Constants.XmlNamespace + "\">" + "<Publisher>Microsoft Corporation</Publisher>" + "</Product>";
                workbook.AddCustomXmlPart(content, Common.Constants.XmlNamespace);
                workbook.AddCustomXmlPart(moreContent, Common.Constants.XmlNamespace);
                string existingContent = workbook.GetCustomXmlPart(Common.Constants.XmlNamespace);
                Assert.AreEqual(moreContent, existingContent);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Verify null output on null inputs
        /// </summary>
        [TestMethod()]
        public void CreateNamedRangeTestNullArguments()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {   
                InteropExcel.Workbook workbook = excelApp.Workbooks.Add();
                string name = null;
                Range range = null;
                Name expected = null;
                Name actual;
                actual = workbook.CreateNamedRange(name, range);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Verify that a named range cannot consist of whitespace only.
        /// </summary>
        [TestMethod()]
        [ExpectedException(typeof(CustomException))]
        public void CreateNamedRangeTestWhitespaceArgument()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                InteropExcel.Workbook workbook = excelApp.Workbooks.Add();
                InteropExcel.Worksheet worksheet = excelApp.ActiveSheet;
                string name = " ";
                Range range = worksheet.get_Range("A1", Type.Missing);
                Name expected = null;
                Name actual;
                actual = WorkbookExtensions.CreateNamedRange(workbook, name, range);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// Verify that an empty string is returned when there are no 
        /// custom xml parts in the workbook
        /// </summary>
        [TestMethod()]
        public void GetCustomXmlPartTestNoParts()
        {   
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {
                InteropExcel.Workbook workbook = excelApp.Workbooks.Add();
                string expected = string.Empty;
                string actual;
                actual = WorkbookExtensions.GetCustomXmlPart(workbook, Common.Constants.XmlNamespace);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// A test for GetSelectionRangeName
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetSelectionRangeNameTest()
        {
            InteropExcel.Application application = new InteropExcel.Application();

            try
            {
                InteropExcel.Workbook workbook = application.OpenWorkbook("TestData.xlsx", false);

                // Get the target range that will be used to set the active sheet
                InteropExcel.Name targetName = workbook.Names.GetNamedRange("GetSelectionRangeName_2");

                // Activate the worksheet which contains the named range.
                ((_Worksheet)targetName.RefersToRange.Worksheet).Activate();

                string expected = "GetSelectionRangeName_3";
                string actual;
                actual = WorkbookExtensions.GetSelectionRangeName(workbook);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for GetSelectionRangeName
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetSelectionRangeNameWithZeroIndexTest()
        {
            InteropExcel.Application application = new InteropExcel.Application();

            try
            {
                InteropExcel.Workbook workbook = application.OpenWorkbook("TestData.xlsx", false);

                // Get the target range that will be used to set the active sheet
                InteropExcel.Name targetName = workbook.Names.GetNamedRange("TestRangeOne");

                // Activate the worksheet which contains the named range.
                ((_Worksheet)targetName.RefersToRange.Worksheet).Activate();

                string expected = string.Format(CultureInfo.InvariantCulture, "{0}_{1}", ((_Worksheet)targetName.RefersToRange.Worksheet).Name, "1");
                string actual;
                actual = WorkbookExtensions.GetSelectionRangeName(workbook);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for GetValidName
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetValidNameTest()
        {
            InteropExcel.Application application = new InteropExcel.Application();

            try
            {
                InteropExcel.Workbook workbook = application.OpenWorkbook("TestData.xlsx", false);

                // Get the target range that will be used to set the active sheet
                InteropExcel.Name targetName = workbook.Names.GetNamedRange("GetValidName");

                // Activate the worksheet which contains the named range.
                ((_Worksheet)targetName.RefersToRange.Worksheet).Activate();

                string name = ((_Worksheet)targetName.RefersToRange.Worksheet).Name;
                string expected = "WWTLayer_012TestS";
                string actual;
                actual = WorkbookExtensions_Accessor.GetValidName(name);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                application.Close();
            }
        }
    }
}
