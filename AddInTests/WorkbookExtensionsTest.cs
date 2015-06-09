//-----------------------------------------------------------------------
// <copyright file="WorkbookExtensionsTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using Microsoft.Office.Interop.Excel;
using Microsoft.Research.Wwt.Excel.Addin;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{    
    /// <summary>
    /// This is a test class for WorkbookExtensionsTest and is intended
    /// to contain all WorkbookExtensionsTest Unit Tests
    /// </summary>
    [TestClass()]
    public class WorkbookExtensionsTest
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
        /// A test for GetViewpointMap
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetViewpointMapTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                ViewpointMap actual = book.GetViewpointMap();
                Assert.AreEqual(actual.Workbook, book);
                Assert.AreEqual(actual.SerializablePerspective.Count, 1);
                Assert.AreEqual(actual.SerializablePerspective[0].Name, "Viewpoint");
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for CleanLayerMap
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void CleanLayerMapTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);

                // Where ever we are creating object for WorkflowController_Accessor, we need to set the ThisAddIn_Accessor.ExcelApplication to 
                // wither null or actual application object.
                ThisAddIn_Accessor.ExcelApplication = application;

                using (WorkflowController_Accessor target = new WorkflowController_Accessor())
                {
                    Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                    Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");
                    target.OnNewWorkbook(book);

                    Name namedRange = book.Names.GetNamedRange("TestProperties_1");
                    namedRange.RefersTo = null;

                    // CurrentWorkbookMap cannot be accessed directly through WorkflowController_Accessor object.
                    WorkbookMap currentWorkbookMap = target.currentWorkbookMap.Target as WorkbookMap;

                    // Before making call to CleanLayerMap method, AllLayerMaps count should be 1.
                    Assert.AreEqual(currentWorkbookMap.AllLayerMaps.Count, 2);

                    Addin.WorkbookExtensions_Accessor.CleanLayerMap(book, target.currentWorkbookMap);

                    // After the call to CleanLayerMap method, AllLayerMaps count should be 0.
                    Assert.AreEqual(currentWorkbookMap.AllLayerMaps.Count, 1);
                }
            }
            finally
            {
                application.Close();
            }
        }
    }
}
