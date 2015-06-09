//-----------------------------------------------------------------------
// <copyright file="TickConverterTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using System;
using System.Collections.ObjectModel;
using System.Globalization;
using Microsoft.Research.Wwt.Excel.Addin;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{    
    /// <summary>
    /// This is a test class for TickConverterTest and is intended
    /// to contain all TickConverterTest Unit Tests
    /// </summary>
    [TestClass()]
    public class TickConverterTest
    {
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
        /// A test for Convert
        /// </summary>
        [TestMethod()]
        public void ConvertTest()
        {
            TickConverter target = new TickConverter();
            object[] values = new object[2];
            Collection<double> doubleValues = new Collection<double>();
            for (double value = 0; value <= 100; value++)
            {
                doubleValues.Add(value * 0.01);
            }
            values[0] = 46;
            values[1] = doubleValues;
            Type targetType = typeof(int);
            object parameter = null;
            CultureInfo culture = CultureInfo.CurrentCulture;
            object expected = "0.45";
            object actual;
            actual = target.Convert(values, targetType, parameter, culture);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for Convert with null arguments
        /// </summary>
        [TestMethod()]
        public void ConvertNullTest()
        {
            TickConverter target = new TickConverter();
            object[] values = null;
            Assert.IsNull(target.Convert(values, typeof(int), null, CultureInfo.CurrentCulture));
        }

        /// <summary>
        /// A test for ConvertBack
        /// </summary>
        [TestMethod()]
        [ExpectedException(typeof(NotImplementedException))]
        public void ConvertBackTest()
        {
            TickConverter target = new TickConverter();
            object value = null; 
            Type[] targetTypes = new Type[] { typeof(double) };
            object parameter = null;
            CultureInfo culture = CultureInfo.CurrentCulture;
            target.ConvertBack(value, targetTypes, parameter, culture);
        }
    }
}
