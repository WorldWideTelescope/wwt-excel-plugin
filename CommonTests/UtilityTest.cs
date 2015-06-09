//-----------------------------------------------------------------------
// <copyright file="UtilityTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Globalization;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// This is a test class for UtilityTest and is intended to contain all UtilityTest Unit Tests
    /// </summary>
    [TestClass()]
    public class UtilityTest
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
        /// A test for SetWWTApplicationPath
        /// </summary>
        [TestMethod()]
        public void SetWWTApplicationPathTest()
        {
            Common.Globals.WWTApplicationPath = string.Empty;
            string expected = string.Empty;

            // Get the application path using the utility method.
            Utility.SetWWTApplicationPath();

            // Get the WWT application path from registry.
            using (RegistryKey registryKey = Registry.ClassesRoot.OpenSubKey(@"WorldWideTelescope.wtml\shell\open\command", false))
            {
                if (registryKey != null)
                {
                    expected = Convert.ToString(registryKey.GetValue(string.Empty), CultureInfo.CurrentCulture);
                    expected = expected.ToString().Split(new char[] { '\"' })[1];
                }
            }

            Assert.AreEqual(Common.Globals.WWTApplicationPath, expected);
        }

        /// <summary>
        /// A test for IsWWTInstalled
        /// </summary>
        [TestMethod()]
        public void IsWWTInstalledTest()
        {
            bool createdWtmlKey = false;

            // Check if WWT is installed or not using registry.
            using (RegistryKey registryKey = Registry.ClassesRoot.OpenSubKey(@".wtml", false))
            {
                if (registryKey == null)
                {
                    // If WWT is really not installed on the machine where unit test is running, create the key so that 
                    // utility method will treat as WWT is installed.
                    using (RegistryKey newRegistryKey = Registry.ClassesRoot.CreateSubKey(@".wtml"))
                    {
                        if (newRegistryKey != null)
                        {
                            createdWtmlKey = true;
                        }
                    }
                }
            }

            try
            {
                Common.Globals.TargetMachine = new TargetMachine();
                Utility.IsWWTInstalled();
            }
            catch (CustomException)
            {
                Assert.Fail("WWT is installed, but utility method says WWT is not installed.");
            }
            finally
            {
                if (createdWtmlKey)
                {
                    // Delete the registry key if it is created by this test case.
                    Registry.ClassesRoot.DeleteSubKey(@".wtml", false);
                }
            }
        }
    }
}