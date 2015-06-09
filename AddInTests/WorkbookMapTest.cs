//-----------------------------------------------------------------------
// <copyright file="WorkbookMapTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Runtime.Serialization;
using Microsoft.Office.Interop.Excel;
using Microsoft.Research.Wwt.Excel.Addin;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{
    /// <summary>
    /// This is a test class for WorkbookMapTest and is intended
    /// to contain all WorkbookMapTest Unit Tests
    /// </summary>
    [TestClass()]
    public class WorkbookMapTest
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
        /// A test for OnSerializingMethod
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnSerializingMethodTest()
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
                    target.OnWorkbookOpen(book);

                    StreamingContext context = new StreamingContext();
                    target.currentWorkbookMap.OnSerializingMethod(context);
                    Assert.IsNotNull(target.currentWorkbookMap.SerializableSelectedLayerMap);
                    Assert.AreEqual(target.currentWorkbookMap.SerializableSelectedLayerMap, target.currentWorkbookMap.SelectedLayerMap);
                }
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for OnSerializingMethod Negative scenario
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnSerializingMethodNegativeTest()
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
                    target.OnWorkbookOpen(book);

                    target.currentWorkbookMap.SelectedLayerMap.RangeName.RefersTo = null;

                    Assert.IsNotNull(target.currentWorkbookMap.SelectedLayerMap);
                    StreamingContext context = new StreamingContext();
                    target.currentWorkbookMap.OnSerializingMethod(context);
                    Assert.IsNull(target.currentWorkbookMap.SelectedLayerMap);
                }
            }
            finally
            {
                application.Close();
            }
        }
    }
}
