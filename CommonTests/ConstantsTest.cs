//-----------------------------------------------------------------------
// <copyright file="ConstantsTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// This is a test class for ConstantsTest and is intended to contain all ConstantsTest Unit Tests
    /// </summary>
    [TestClass()]
    public class ConstantsTest
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
        /// A test for DefaultStartTime
        /// </summary>
        [TestMethod()]
        public void DefaultStartTimeTest()
        {
            DateTime actual = Constants.DefaultStartTime;
            DateTime expected = DateTime.MinValue;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for DefaultEndTime
        /// </summary>
        [TestMethod()]
        public void DefaultEndTimeTest()
        {
            DateTime actual = Constants.DefaultEndTime;
            DateTime expected = DateTime.MaxValue;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for DefaultErrorResponse
        /// </summary>
        [TestMethod()]
        public void DefaultErrorResponseTest()
        {
            string actual = Constants.DefaultErrorResponse;
            string expected = "<LayerApi><Status>Error</Status></LayerApi>";
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for DefaultFadeSpan
        /// </summary>
        [TestMethod()]
        public void DefaultFadeSpanTest()
        {
            TimeSpan actual = Constants.DefaultFadeSpan;
            TimeSpan expected = new TimeSpan();
            Assert.AreEqual(expected, actual);
        }
    }
}
