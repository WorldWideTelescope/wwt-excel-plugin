//-----------------------------------------------------------------------
// <copyright file="UpdateManagerTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using Microsoft.Research.Wwt.Excel.Addin;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{
    /// <summary>
    /// This is a test class for UpdateManagerTest and is intended
    /// to contain all UpdateManagerTest Unit Tests
    /// </summary>
    [TestClass()]
    public class UpdateManagerTest
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
        /// A test for UpdateManager Constructor
        /// </summary>
        [TestMethod()]
        public void UpdateManagerConstructorTest()
        {
            using (UpdateManager_Accessor target = new UpdateManager_Accessor())
            {
                Assert.IsNotNull(target);
                Assert.IsNotNull(target.checkUpdatesWorker);
                Assert.IsNotNull(target.downloadLinkWorker);
                Assert.IsNotNull(target.installUpdateWorker);
            }
        }
    }
}
