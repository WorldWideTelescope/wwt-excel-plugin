//-----------------------------------------------------------------------
// <copyright file="StringExtensionsTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Globalization;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// This is a test class for StringExtensionsTest and is intended
    /// to contain all StringExtensionsTest Unit Tests
    /// </summary>
    [TestClass()]
    public class StringExtensionsTest
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
        /// A test for AsBoolean
        /// </summary>
        [TestMethod()]
        public void AsBooleanTest()
        {
            bool defaultValue = false;
            bool expected = true;
            string value = Convert.ToString(expected, CultureInfo.InvariantCulture);
            bool actual = StringExtensions.AsBoolean(value, defaultValue);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for AsBoolean
        /// </summary>
        [TestMethod()]
        public void AsBooleanNegativeTest()
        {
            string value = "Test";
            bool defaultValue = false;
            bool actual = value.AsBoolean(defaultValue);
            Assert.AreEqual(defaultValue, actual);
        }

        /// <summary>
        /// A test for AsDateTime
        /// </summary>
        [TestMethod()]
        public void AsDateTimeTest()
        {
            string value = Convert.ToString(DateTime.Now, CultureInfo.InvariantCulture);
            DateTime expected = DateTime.Parse(value, CultureInfo.InvariantCulture);
            DateTime defaultValue = new DateTime();
            DateTime actual = StringExtensions.AsDateTime(value, defaultValue);
            Assert.IsTrue(actual.Equals(expected));
        }

        /// <summary>
        /// A test for AsDateTime
        /// </summary>
        [TestMethod()]
        public void AsDateTimeNegativeTest()
        {
            string value = "02/79/11";
            DateTime defaultValue = new DateTime();
            DateTime actual = StringExtensions.AsDateTime(value, defaultValue);
            Assert.IsTrue(actual.Equals(defaultValue));
        }

        /// <summary>
        /// A test for AsDouble
        /// </summary>
        [TestMethod()]
        public void AsDoubleTest()
        {
            double defaultValue = 0F;
            double expected = 5F;
            string value = Convert.ToString(expected, CultureInfo.InvariantCulture);
            double actual = StringExtensions.AsDouble(value, defaultValue);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for AsDouble
        /// </summary>
        [TestMethod()]
        public void AsDoubleNegativeTest()
        {
            double defaultValue = 0F;
            string value = "a12";
            double actual = StringExtensions.AsDouble(value, defaultValue);
            Assert.AreEqual(defaultValue, actual);
        }

        /// <summary>
        /// A test for AsEnum
        /// </summary>
        [TestMethod()]
        public void AsEnumTest()
        {
            AltType defaultValue = AltType.Altitude;
            AltType expected = AltType.Depth;
            string value = Convert.ToString(expected, CultureInfo.InvariantCulture);
            AltType actual = StringExtensions.AsEnum<AltType>(value, defaultValue);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for AsEnum
        /// </summary>
        [TestMethod()]
        public void AsEnumNegativeTest()
        {
            AltType defaultValue = AltType.Altitude;
            string value = "Magnitude";
            AltType actual = StringExtensions.AsEnum<AltType>(value, defaultValue);
            Assert.AreEqual(defaultValue, actual);
        }

        /// <summary>
        /// A test for AsInteger
        /// </summary>
        [TestMethod()]
        public void AsIntegerTest()
        {
            int defaultValue = 0;
            int expected = 45;
            string value = Convert.ToString(expected, CultureInfo.InvariantCulture);
            int actual = StringExtensions.AsInteger(value, defaultValue);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for AsInteger
        /// </summary>
        [TestMethod()]
        public void AsIntegerNegativeTest()
        {
            int defaultValue = 0;
            string value = "a1";
            int actual = StringExtensions.AsInteger(value, defaultValue);
            Assert.AreEqual(defaultValue, actual);
        }

        /// <summary>
        /// A test for AsTimeSpan
        /// </summary>
        [TestMethod()]
        public void AsTimeSpanTest()
        {
            TimeSpan defaultValue = new TimeSpan(0, 0, 0);
            TimeSpan expected = new TimeSpan(2, 3, 4);
            string value = Convert.ToString(expected, CultureInfo.InvariantCulture);
            TimeSpan actual = StringExtensions.AsTimeSpan(value, defaultValue);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for AsTimeSpan
        /// </summary>
        [TestMethod()]
        public void AsTimeSpanNegativeTest()
        {
            TimeSpan defaultValue = new TimeSpan(0, 0, 0);
            string value = "123456789";
            TimeSpan actual = StringExtensions.AsTimeSpan(value, defaultValue);
            Assert.AreEqual(defaultValue, actual);
        }
    }
}
