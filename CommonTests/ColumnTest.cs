//-----------------------------------------------------------------------
// <copyright file="ColumnTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Collections.ObjectModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{   
    /// <summary>
    /// This is a test class for ColumnTest and is intended
    /// to contain all ColumnTest Unit Tests
    /// </summary>
    [TestClass()]
    public class ColumnTest
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
        /// A test for ColumnDisplayValue
        /// </summary>
        [TestMethod()]
        public void ColumnDisplayValueTest()
        {
            ColumnType columnType = ColumnType.Lat; 
            string columnDisplayValue = "4.56"; 
            Collection<string> columnComparisonValue = new Collection<string>();
            columnComparisonValue.Add(columnDisplayValue);
            Column target = new Column(columnType, columnDisplayValue, columnComparisonValue);
            string expected = columnDisplayValue;
            target.ColumnDisplayValue = expected;
            string actual = target.ColumnDisplayValue;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for ColType
        /// </summary>
        [TestMethod()]
        public void ColTypeTest()
        {
            string columnDisplayValue = "Latitude";
            Collection<string> columnComparisonValue = new Collection<string> { "LAT" };
            Column target = new Column(ColumnType.Lat, columnDisplayValue, columnComparisonValue);

            ColumnType expected = ColumnType.Lat;
            ColumnType actual = target.ColType;

            Assert.AreEqual(expected, actual);

            target.ColType = ColumnType.Geo;
            expected = ColumnType.Geo;
            actual = target.ColType;

            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for ColumnMatchValues
        /// </summary>
        [TestMethod()]
        public void ColumnMatchValuesTest()
        {
            string columnDisplayValue = "Latitude";
            Collection<string> columnComparisonValue = new Collection<string> { "LAT" };
            Column target = new Column(ColumnType.Lat, columnDisplayValue, columnComparisonValue);

            Collection<string> expected = new Collection<string> { "LAT" };

            ReadOnlyCollection<string> actual = target.ColumnMatchValues;

            Assert.AreEqual(expected.Count, actual.Count);
            Assert.AreEqual(expected[0], actual[0]);
        }

        /// <summary>
        /// A test for Column Constructor
        /// </summary>
        [TestMethod()]
        public void ColumnConstructorTest()
        {
            string columnDisplayValue = "Longitude";
            Collection<string> columnComparisonValue = new Collection<string> { "LON" };
            Column target = new Column(ColumnType.Long, columnDisplayValue, columnComparisonValue);

            Assert.IsNotNull(target);
            Assert.AreEqual(target.ColType, ColumnType.Long);
            Assert.AreEqual(target.ColumnDisplayValue, "Longitude");
            Assert.AreEqual(target.ColumnMatchValues.Count, columnComparisonValue.Count);
            Assert.AreEqual(target.ColumnMatchValues[0], columnComparisonValue[0]);
        }
    }
}
