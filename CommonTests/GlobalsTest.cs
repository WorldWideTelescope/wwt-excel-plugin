//-----------------------------------------------------------------------
// <copyright file="GlobalsTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// This is a test class for GlobalsTest and is intended to contain all GlobalsTest Unit Tests
    /// </summary>
    [TestClass()]
    public class GlobalsTest
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
        /// A test for AddinTraceSource
        /// </summary>
        [TestMethod()]
        public void AddinTraceSourceTest()
        {
            TraceSource actual = Globals.AddinTraceSource;
            Assert.IsNotNull(actual);
            Assert.AreEqual(actual.Name, "WWTEarthAddin");
        }
    }
}
