//-----------------------------------------------------------------------
// <copyright file="LayerMapTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using System;
using System.Collections.ObjectModel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Research.Wwt.Excel.Addin;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using InteropExcel = Microsoft.Office.Interop.Excel;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{
    /// <summary>
    /// This is a test class for LayerMapTest and is intended to contain all LayerMapTest Unit Tests
    /// </summary>
    [TestClass()]
    public class LayerMapTest
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
        /// A test for ColumnsList
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void ColumnsListTest()
        {
            InteropExcel.Application application = new InteropExcel.Application();

            try
            {
                InteropExcel.Workbook book = application.OpenWorkbook("TestData.xlsx", false);

                // Get the named range stored in the test data excel file.
                Name name = book.Names.GetNamedRange("ColumnList");
                LayerMap target = new LayerMap(name);

                target.ColumnsList = ColumnExtensions.PopulateColumnList();

                // Check the count of columns.
                Assert.AreEqual(target.ColumnsList.Count, 19);

                // Check the column names
                Assert.AreEqual(target.ColumnsList[0].ColumnDisplayValue, "Select One");
                Assert.AreEqual(target.ColumnsList[1].ColumnDisplayValue, "Latitude");
                Assert.AreEqual(target.ColumnsList[2].ColumnDisplayValue, "Longitude");
                Assert.AreEqual(target.ColumnsList[3].ColumnDisplayValue, "Start Date");
                Assert.AreEqual(target.ColumnsList[4].ColumnDisplayValue, "End Date");
                Assert.AreEqual(target.ColumnsList[5].ColumnDisplayValue, "Depth");
                Assert.AreEqual(target.ColumnsList[6].ColumnDisplayValue, "Altitude");
                Assert.AreEqual(target.ColumnsList[7].ColumnDisplayValue, "Distance");
                Assert.AreEqual(target.ColumnsList[8].ColumnDisplayValue, "Magnitude");
                Assert.AreEqual(target.ColumnsList[9].ColumnDisplayValue, "Geometry");
                Assert.AreEqual(target.ColumnsList[10].ColumnDisplayValue, "Color");
                Assert.AreEqual(target.ColumnsList[11].ColumnDisplayValue, "RA");
                Assert.AreEqual(target.ColumnsList[12].ColumnDisplayValue, "Dec");
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for ResetRange
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void ResetRangeTest()
        {
            InteropExcel.Application application = new InteropExcel.Application();

            try
            {
                InteropExcel.Workbook book = application.OpenWorkbook("TestData.xlsx", false);
                Name name = book.Names.GetNamedRange("TestRangeOne");
                LayerMap target = new LayerMap(name);
                Name resetRangeName = book.Names.GetNamedRange("TestRangeTarget");
                string expected = resetRangeName.Name;
                target.ResetRange(resetRangeName);
                string actual = target.RangeDisplayName;
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for SetMappingOnSizeColumn
        /// </summary>
        [TestMethod()]
        public void SetMappingOnSizeColumnTest()
        {
            Collection<ColumnType> mappedColTypes = new Collection<ColumnType>();
            mappedColTypes.Add(ColumnType.RA);
            mappedColTypes.Add(ColumnType.Dec);
            mappedColTypes.Add(ColumnType.None);
            mappedColTypes.Add(ColumnType.Long);

            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.StartTime = DateTime.Now;
            layer.SizeColumn = 2;

            LayerMap_Accessor layerMapAccessor = new LayerMap_Accessor(layer);
            layerMapAccessor.MappedColumnType = mappedColTypes;

            Assert.AreEqual(mappedColTypes[2], ColumnType.None);
            layerMapAccessor.SetMappingOnSizeColumn();
            Assert.AreEqual(mappedColTypes[2], ColumnType.Mag);
        }
    }
}