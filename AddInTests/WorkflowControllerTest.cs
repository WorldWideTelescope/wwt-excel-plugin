//-----------------------------------------------------------------------
// <copyright file="WorkflowControllerTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Research.Wwt.Excel.Addin;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{
    /// <summary>
    /// This is a test class for WorkflowControllerTest and is intended to contain all WorkflowControllerTest Unit Tests
    /// </summary>
    [TestClass()]
    public class WorkflowControllerTest
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
        /// A test for GetActiveWorksheet
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetActiveWorksheetTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);

                ThisAddIn_Accessor.ExcelApplication = application;

                book.Sheets[3].Activate();

                // Active sheet returned will be the book.Sheets[3] object. 
                _Worksheet activeSheet = WorkflowController_Accessor.GetActiveWorksheet();
                Assert.AreEqual(activeSheet, book.Sheets[3]);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for GetRangeDifference
        /// </summary>
        [TestMethod()]
        public void GetRangeDifferenceTest()
        {
            int rangeLength = 10;
            int dataLength = 8;
            int expected = 2;

            // Difference should be 2.
            int actual = WorkflowController_Accessor.GetRangeDifference(rangeLength, dataLength);
            Assert.AreEqual(expected, actual);

            rangeLength = 8;
            dataLength = 10;

            // Even if rangeLength is less than dataLength, difference should be 2, not -2.
            actual = WorkflowController_Accessor.GetRangeDifference(rangeLength, dataLength);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for SetGetLayerDataDisplayName
        /// </summary>
        [TestMethod()]
        public void SetGetLayerDataDisplayNameTest()
        {
            WorkflowController_Accessor target = GetWorkflowControllerAccessor(null);
            Layer layer = new Layer();

            LayerMap_Accessor selectedLayerMap = new LayerMap_Accessor(layer);
            target.layerDetailsViewModel = new LayerDetailsViewModel();

            // When MapType is Local/WWT, LayerDataDisplayName will be "Get Layer Data"; 
            string actual = "Get Layer Data";
            selectedLayerMap.MapType = LayerMapType.Local;
            target.SetGetLayerDataDisplayName(selectedLayerMap);

            string expected = target.layerDetailsViewModel.LayerDataDisplayName;

            Assert.AreEqual(actual, expected);

            // When MapType is LocalInWWT, LayerDataDisplayName will be "Refresh"; 
            actual = "Refresh";
            selectedLayerMap.MapType = LayerMapType.LocalInWWT;
            target.SetGetLayerDataDisplayName(selectedLayerMap);

            expected = target.layerDetailsViewModel.LayerDataDisplayName;

            Assert.AreEqual(actual, expected);
        }

        /// <summary>
        /// A test for InsertRows
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void InsertRowsTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);

                // Get the named range stored in the test data excel file.
                // This range contains first two rows referred with address "$A$1:$E$2".
                Name name = book.Names.GetNamedRange("InsertRows");

                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);

                target.InsertRows(name.RefersToRange);

                // After two rows are inserted, address of the named range will change to "$A$3:$E$4".
                Assert.AreEqual("$A$3:$E$4", name.RefersToRange.Address);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for InsertColumns
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void InsertColumnsTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);

                // Get the named range stored in the test data excel file.
                // This range contains first two columns referred with address "$A$1:$B$8".
                Name name = book.Names.GetNamedRange("InsertColumns");

                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);

                target.InsertColumns(name.RefersToRange);

                // After two columns are inserted, address of the named range will change to "$C$1:$D$8".
                Assert.AreEqual("$C$1:$D$8", name.RefersToRange.Address);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for GetSelectedLayerWorksheet
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetSelectedLayerWorksheetTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);

                // Get the named range stored in the test data excel file.
                // This range refers to address "$A$1:$D$7".
                Name name = book.Names.GetNamedRange("GetSelectedLayerWorksheet");

                ThisAddIn_Accessor.ExcelApplication = application;
                LayerMap_Accessor selectedLayerMap = new LayerMap_Accessor(name);

                // Active the first sheet, to verify the actual sheet to which the range belongs to is getting selected.
                book.Sheets[1].Activate();

                // GetSelectedLayerWorksheet range belongs to sheet 6.
                _Worksheet expected = book.Sheets[6];
                _Worksheet actual = WorkflowController_Accessor.GetSelectedLayerWorksheet(selectedLayerMap);

                // Verify that both Worksheet objects are same.
                Assert.AreEqual(expected, actual);

                // Also verify that the active sheet is same as the one returned.
                Assert.AreEqual(application.ActiveSheet, actual);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for CreateRangeForLayer
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void CreateRangeForLayerTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);

                // Get the named range stored in the test data excel file.
                // This range refers to address "A$1:$D$8" in sheet CreateRangeForLayer.
                Name namedRange = book.Names.GetNamedRange("CreateRangeForLayer");

                // Activate the worksheet which contains the named range.
                ((_Worksheet)namedRange.RefersToRange.Worksheet).Activate();

                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);

                target.CreateRangeForLayer(namedRange.RefersToRange);

                Name newRange = book.Names.GetNamedRange("CreateRangeForLayer_1");

                if (null == newRange)
                {
                    Assert.Fail("New range for layer is not created successfully!");
                }
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for BuildReferenceFrameDropDown
        /// </summary>
        [TestMethod()]
        public void BuildReferenceFrameDropDownTest()
        {
            WorkflowController_Accessor target = GetWorkflowControllerAccessor(null);
            target.layerDetailsViewModel = new LayerDetailsViewModel();
            Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
            Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

            ICollection<Group> groups = WWTManager_Accessor.GetAllWWTGroups(true);

            target.BuildReferenceFrameDropDown();
            Assert.AreEqual(1, target.layerDetailsViewModel.ReferenceGroups.Count);
            foreach (GroupViewModel groupView in target.layerDetailsViewModel.ReferenceGroups)
            {
                Assert.AreEqual(groupView.Name, string.Empty);
                int index = 0;
                foreach (Group group in groups)
                {
                    Assert.AreEqual(group.Name, groupView.ReferenceGroup[index].Name);
                    Assert.AreEqual(group.Path, groupView.ReferenceGroup[index].Path);
                    index++;
                }
            }
            Assert.AreEqual(1, target.layerDetailsViewModel.ReferenceGroups.Count);
        }

        /// <summary>
        /// A test for RebuildGroupLayerDropdown
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void RebuildGroupLayerDropdownTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);

                List<LayerMapDropDownViewModel> layermapViews = new List<LayerMapDropDownViewModel>();
                layermapViews.Add(new LayerMapDropDownViewModel() { ID = "-1", Name = "Select One" });
                layermapViews.Add(new LayerMapDropDownViewModel() { ID = "0", Name = "Local Layers" });
                layermapViews.Add(new LayerMapDropDownViewModel() { ID = "1", Name = "WWT Layers" });

                target.layerDetailsViewModel = new LayerDetailsViewModel();
                target.layerDetailsViewModel.Layers = new System.Collections.ObjectModel.ObservableCollection<LayerMapDropDownViewModel>();
                target.currentWorkbookMap = new WorkbookMap_Accessor(book);
                target.RebuildGroupLayerDropdown();
                int index = 0;
                foreach (LayerMapDropDownViewModel layermapView in target.layerDetailsViewModel.Layers)
                {
                    Assert.AreEqual(layermapView.Name, layermapViews[index].Name);
                    Assert.AreEqual(layermapView.ID, layermapViews[index].ID);
                    index++;
                }
                Assert.AreEqual(3, target.layerDetailsViewModel.Layers.Count);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for UpdateHeader
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void UpdateHeaderTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                target.currentWorkbookMap = new WorkbookMap_Accessor(book);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                // Get the named range stored in the test data excel file.
                Name name = book.Names.GetNamedRange("ColumnList");

                LayerMap_Accessor wwtLayerMap = new LayerMap_Accessor(name);
                wwtLayerMap.LayerDetails.Name = "Layer";
                wwtLayerMap.LayerDetails.ID = "754becad-8e2c-4452-b7e9-55827d4d2786";
                wwtLayerMap.MapType = LayerMapType.WWT;

                WorkflowController_Accessor.UpdateHeader(wwtLayerMap);

                // It is expected to be in synch
                Assert.IsFalse(wwtLayerMap.IsNotInSync);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTInvalidMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                WorkflowController_Accessor.UpdateHeader(wwtLayerMap);

                // It is expected not to be in synch
                Assert.IsTrue(wwtLayerMap.IsNotInSync);

                LayerMap_Accessor localLayerMap = new LayerMap_Accessor(name);
                localLayerMap.LayerDetails.Name = "Layer";
                localLayerMap.LayerDetails.ID = "754becad-8e2c-4452-b7e9-55827d4d2786";
                localLayerMap.MapType = LayerMapType.Local;
                localLayerMap.IsNotInSync = true;

                WorkflowController_Accessor.UpdateHeader(localLayerMap);

                // It is expected not to be in synch
                Assert.IsTrue(localLayerMap.IsNotInSync);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for CreateGroupInWWT
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void CreateGroupInWWTTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                target.currentWorkbookMap = new WorkbookMap_Accessor(book);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                // Get the named range stored in the test data excel file.
                Name name = book.Names.GetNamedRange("ColumnList");
                Group parentGroup = new Group("Sun", GroupType.ReferenceFrame, null);
                Group childGroup = new Group("Earth", GroupType.LayerGroup, parentGroup);

                LayerMap_Accessor wwtLayerMap = new LayerMap_Accessor(name);
                wwtLayerMap.LayerDetails.Name = "Layer";
                wwtLayerMap.LayerDetails.ID = "754becad-8e2c-4452-b7e9-55827d4d2786";
                wwtLayerMap.LayerDetails.Group = childGroup;
                wwtLayerMap.MapType = LayerMapType.WWT;

                bool expected = true;
                bool actual = WorkflowController_Accessor.CreateGroupInWWT(wwtLayerMap);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for CreateLayerInWWT
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void CreateLayerInWWTTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);

                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                target.currentWorkbookMap = new WorkbookMap_Accessor(book);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                // Get the named range stored in the test data excel file.
                Name name = book.Names.GetNamedRange("ColumnList");
                Group parentGroup = new Group("Sun", GroupType.ReferenceFrame, null);
                Group childGroup = new Group("Earth", GroupType.LayerGroup, parentGroup);

                string oldLayerId = "754becad-8e2c-4452-b7e9-55827d4d2786";

                LayerMap_Accessor localLayerMap = new LayerMap_Accessor(name);
                localLayerMap.LayerDetails.Name = "Layer";
                localLayerMap.LayerDetails.ID = oldLayerId;
                localLayerMap.LayerDetails.Group = childGroup;
                localLayerMap.MapType = LayerMapType.Local;
                localLayerMap.HeaderRowData = new System.Collections.ObjectModel.Collection<string>();
                localLayerMap.HeaderRowData.Add("LAT");
                localLayerMap.HeaderRowData.Add("Long");
                localLayerMap.HeaderRowData.Add("Time");

                bool expected = true;

                string layerId = "2cf4374f-e1ce-47a9-b08c-31079765ddcf";
                bool actual = WorkflowController_Accessor.CreateLayerInWWT(localLayerMap);
                Assert.AreEqual(expected, actual);
                Assert.AreEqual(LayerMapType.LocalInWWT, localLayerMap.MapType);
                Assert.AreEqual(layerId, localLayerMap.LayerDetails.ID);

                localLayerMap.LayerDetails.ID = oldLayerId;
                Group invalidGroup = new Group("Test", GroupType.LayerGroup, parentGroup);
                localLayerMap.LayerDetails.Group = invalidGroup;
                actual = WorkflowController_Accessor.CreateLayerInWWT(localLayerMap);
                Assert.AreEqual(expected, actual);
                Assert.AreEqual(LayerMapType.LocalInWWT, localLayerMap.MapType);
                Assert.AreEqual(layerId, localLayerMap.LayerDetails.ID);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for CreateLayerInWWT
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void CreateIfNotExistTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                target.currentWorkbookMap = new WorkbookMap_Accessor(book);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                // Get the named range stored in the test data excel file.
                Name name = book.Names.GetNamedRange("ColumnList");
                Group parentGroup = new Group("Sun", GroupType.ReferenceFrame, null);
                Group childGroup = new Group("Earth", GroupType.LayerGroup, parentGroup);

                string layerId = "2cf4374f-e1ce-47a9-b08c-31079765ddcf";
                string oldLayerId = "754becad-8e2c-4452-b7e9-55827d4d2786";

                LayerMap_Accessor localLayerMap = new LayerMap_Accessor(name);
                localLayerMap.LayerDetails.Name = "Layer";
                localLayerMap.LayerDetails.Group = childGroup;
                localLayerMap.MapType = LayerMapType.Local;
                localLayerMap.HeaderRowData = new System.Collections.ObjectModel.Collection<string>();
                localLayerMap.HeaderRowData.Add("LAT");
                localLayerMap.HeaderRowData.Add("Long");
                localLayerMap.HeaderRowData.Add("Time");

                bool expected = true;
                bool actual = WorkflowController_Accessor.CreateIfNotExist(localLayerMap);
                Assert.AreEqual(expected, actual);
                Assert.AreEqual(layerId, localLayerMap.LayerDetails.ID);

                localLayerMap.MapType = LayerMapType.LocalInWWT;
                localLayerMap.LayerDetails.ID = oldLayerId;
                actual = WorkflowController_Accessor.CreateIfNotExist(localLayerMap);
                Assert.AreEqual(expected, actual);
                Assert.AreEqual(layerId, localLayerMap.LayerDetails.ID);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for GetCurrentRangeLayer
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetCurrentRangeLayerTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");
                target.OnWorkbookOpen(book);
                book.Sheets[12].Activate();
                Worksheet worksheet = (Worksheet)book.Sheets[12];
                Range selectedRange = worksheet.get_Range("A2:H6");
                LayerMap_Accessor layerMap = target.GetCurrentRangeLayer(selectedRange);
                Assert.AreEqual("TestProperties_1", layerMap.RangeDisplayName);
                Assert.AreEqual("TestProperties_1", layerMap.LayerDetails.Name);
            }
            finally
            {
                application.Close();
            }
        }

        #region UpdateDataTest

        /// <summary>
        /// A test for UpdateData
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void UpdateDataTestSuccess()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                // Get the named range stored in the test data excel file.
                // This range refers to address "A$1:$D$8" in sheet CreateRangeForLayer.
                Name namedRange = book.Names.GetNamedRange("GetLastUsedGroupTestLatLon");

                Range selectedRange = namedRange.RefersToRange;
                LayerMap_Accessor selectedlayer = new LayerMap_Accessor(namedRange);
                selectedlayer.MapType = LayerMapType.LocalInWWT;
                selectedlayer.IsNotInSync = false;
                selectedlayer.LayerDetails.ID = "Exits";

                WorkflowController_Accessor.UpdateData(selectedRange, selectedlayer);
                Assert.IsTrue(!selectedlayer.IsNotInSync);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for UpdateData
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void UpdateDataTestFailure()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTInvalidMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                // Get the named range stored in the test data excel file.
                // This range refers to address "A$1:$D$8" in sheet CreateRangeForLayer.
                Name namedRange = book.Names.GetNamedRange("GetLastUsedGroupTestLatLon");

                Range selectedRange = namedRange.RefersToRange;
                LayerMap_Accessor selectedlayer = new LayerMap_Accessor(namedRange);
                selectedlayer.MapType = LayerMapType.LocalInWWT;
                selectedlayer.IsNotInSync = false;
                selectedlayer.LayerDetails.ID = "NotExits";

                WorkflowController_Accessor.UpdateData(selectedRange, selectedlayer);

                Assert.IsTrue(selectedlayer.IsNotInSync);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion

        #region UpdateWWTTest

        /// <summary>
        /// A test for UpdateWWT
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void UpdateWWTTestSuccess()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                // Get the named range stored in the test data excel file.
                // This range refers to address "A$1:$D$8" in sheet CreateRangeForLayer.
                Name namedRange = book.Names.GetNamedRange("GetLastUsedGroupTestLatLon");

                LayerMap_Accessor selectedlayer = new LayerMap_Accessor(namedRange);
                selectedlayer.MapType = LayerMapType.LocalInWWT;
                selectedlayer.IsNotInSync = false;
                selectedlayer.LayerDetails.ID = "Exits";

                bool actual;

                actual = WorkflowController_Accessor.UpdateWWT(selectedlayer);

                Assert.AreEqual(true, actual);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for UpdateWWT
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void UpdateWWTTestFailure()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTInvalidMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                // Get the named range stored in the test data excel file.
                // This range refers to address "A$1:$D$8" in sheet CreateRangeForLayer.
                Name namedRange = book.Names.GetNamedRange("GetLastUsedGroupTestLatLon");

                LayerMap_Accessor selectedlayer = new LayerMap_Accessor(namedRange);
                selectedlayer.MapType = LayerMapType.Local;
                selectedlayer.IsNotInSync = false;
                selectedlayer.LayerDetails.ID = "Exits";

                bool actual = WorkflowController_Accessor.UpdateWWT(selectedlayer);

                Assert.AreEqual(false, actual);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for UpdateWWT
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void UpdateWWTTestNullLayerID()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTInvalidMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                // Get the named range stored in the test data excel file.
                // This range refers to address "A$1:$D$8" in sheet CreateRangeForLayer.
                Name namedRange = book.Names.GetNamedRange("GetLastUsedGroupTestLatLon");

                LayerMap_Accessor selectedlayer = new LayerMap_Accessor(namedRange);
                selectedlayer.MapType = LayerMapType.Local;
                selectedlayer.IsNotInSync = false;
                selectedlayer.LayerDetails.ID = string.Empty;

                bool actual = WorkflowController_Accessor.UpdateWWT(selectedlayer);

                Assert.AreEqual(false, actual);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion

        #region GetLastUsedGroupTest

        /// <summary>
        /// A test for GetLastUsedGroup
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetLastUsedGroupTestForSky()
        {
            Application application = new Application();
            try
            {
                Collection<Group> groups = new Collection<Group>();

                bool isWWTRunning = false;
                Group expected = groups.GetDefaultSkyGroup();

                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                // Get the named range stored in the test data excel file.
                Name namedRange = book.Names.GetNamedRange("GetLastUsedGroupTestRaDec");

                LayerMap_Accessor selectedlayer = new LayerMap_Accessor(namedRange);
                selectedlayer.MapType = LayerMapType.Local;
                selectedlayer.IsNotInSync = false;
                selectedlayer.LayerDetails.ID = "Exits";

                Group actual = WorkflowController_Accessor.GetLastUsedGroup(selectedlayer, isWWTRunning);
                Assert.AreEqual(expected.Name, actual.Name);
                Assert.AreEqual(expected.Path, actual.Path);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for GetLastUsedGroup
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetLastUsedGroupTestForEarth()
        {
            Application application = new Application();
            try
            {
                Collection<Group> groups = new Collection<Group>();

                bool isWWTRunning = false;
                Group expected = groups.GetDefaultEarthGroup();

                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                // Get the named range stored in the test data excel file.
                Name namedRange = book.Names.GetNamedRange("SetAllPropertiesRange");

                LayerMap_Accessor selectedlayer = new LayerMap_Accessor(namedRange);
                selectedlayer.MapType = LayerMapType.Local;
                selectedlayer.IsNotInSync = false;
                selectedlayer.LayerDetails.ID = "Exits";

                Group actual = WorkflowController_Accessor.GetLastUsedGroup(selectedlayer, isWWTRunning);
                Assert.AreEqual(expected.Name, actual.Name);
                Assert.AreEqual(expected.Path, actual.Path);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for GetLastUsedGroup
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetLastUsedGroupTestForLastUsedGroupNotPresent()
        {
            Application application = new Application();
            try
            {
                Collection<Group> groups = new Collection<Group>();

                bool isWWTRunning = false;
                Group expected = groups.GetDefaultEarthGroup();

                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                // Get the named range stored in the test data excel file.
                Name namedRange = book.Names.GetNamedRange("GetLastUsedGroupTestMap");

                LayerMap_Accessor selectedlayer = new LayerMap_Accessor(namedRange);
                selectedlayer.MapType = LayerMapType.Local;
                selectedlayer.IsNotInSync = false;
                selectedlayer.LayerDetails.ID = "Exits";

                Group actual = WorkflowController_Accessor.GetLastUsedGroup(selectedlayer, isWWTRunning);
                Assert.AreEqual(expected.Name, actual.Name);
                Assert.AreEqual(expected.Path, actual.Path);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for GetLastUsedGroup
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetLastUsedGroupTestForLastUsedGroupNotPresentWWTRunning()
        {
            Application application = new Application();
            try
            {
                Collection<Group> groups = new Collection<Group>();

                bool isWWTRunning = true;
                Group expected = groups.GetDefaultEarthGroup();

                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                // Get the named range stored in the test data excel file.
                Name namedRange = book.Names.GetNamedRange("GetLastUsedGroupTestMap");

                LayerMap_Accessor selectedlayer = new LayerMap_Accessor(namedRange);
                selectedlayer.MapType = LayerMapType.Local;
                selectedlayer.IsNotInSync = false;
                selectedlayer.LayerDetails.ID = "Exits";

                // Make sure LastUsedGroup object is null
                WorkflowController_Accessor.LastUsedGroup = null;

                Group actual = WorkflowController_Accessor.GetLastUsedGroup(selectedlayer, isWWTRunning);
                Assert.AreEqual(expected.Name, actual.Name);
                Assert.AreEqual(expected.Path, actual.Path);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for GetLastUsedGroup
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetLastUsedGroupTestForLastUsedGroupPresentAndDeleted()
        {
            Application application = new Application();
            try
            {
                Collection<Group> groups = new Collection<Group>();

                bool isWWTRunning = false;
                Group expected = groups.GetDefaultEarthGroup();

                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                // Get the named range stored in the test data excel file.
                Name namedRange = book.Names.GetNamedRange("GetLastUsedGroupTestMap");

                LayerMap_Accessor selectedlayer = new LayerMap_Accessor(namedRange);
                selectedlayer.MapType = LayerMapType.Local;
                selectedlayer.IsNotInSync = false;
                selectedlayer.LayerDetails.ID = "Exits";

                WorkflowController_Accessor.LastUsedGroup = new Group("sky", GroupType.ReferenceFrame, null) { IsDeleted = true };

                Group actual = WorkflowController_Accessor.GetLastUsedGroup(selectedlayer, isWWTRunning);
                Assert.AreEqual(expected.Name, actual.Name);
                Assert.AreEqual(expected.Path, actual.Path);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion

        #region GotoViewpointTest

        /// <summary>
        /// A test for GotoViewpoint
        /// </summary>
        [TestMethod()]
        public void GotoViewpointTest()
        {
            bool createdWtmlKey = false;
            try
            {
                createdWtmlKey = CreateWwtKeyIfNotExists();
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                Perspective perspective = new Perspective(Common.Constants.EarthLookAt, Common.Constants.EarthLookAt, false, Common.Constants.LatitudeDefaultValue, Common.Constants.LongitudeDefaultValue, Common.Constants.ZoomDefaultValue, Common.Constants.RotationDefaultValue, Common.Constants.LookAngleDefaultValue, DateTime.Now.ToString(), Common.Constants.TimeRateDefaultValue, Common.Constants.EarthZoomTextDefaultValue, string.Empty);

                WorkflowController_Accessor.GotoViewpoint(perspective);
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

        #endregion

        #region GotoViewpointFromDataTest

        /// <summary>
        /// A test for GotoViewpointFromData
        /// </summary>
        [TestMethod()]
        public void GotoViewpointFromDataTest()
        {
            bool createdWtmlKey = false;
            try
            {
                createdWtmlKey = CreateWwtKeyIfNotExists();
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                Perspective perspective = new Perspective(Common.Constants.SkyLookAt, Common.Constants.SkyReferenceFrame, true, Common.Constants.LatitudeDefaultValue, Common.Constants.LongitudeDefaultValue, Common.Constants.ZoomDefaultValue, Common.Constants.RotationDefaultValue, Common.Constants.LookAngleDefaultValue, DateTime.Now.ToString(), Common.Constants.TimeRateDefaultValue, Common.Constants.SkyZoomTextDefaultValue, string.Empty);

                WorkflowController_Accessor.GotoViewpointFromData(perspective);
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

        #endregion

        #region GotoViewpointOnViewInWWTTest

        /// <summary>
        /// A test for GotoViewpointOnViewInWWT
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GotoViewpointOnViewInWWTTestEarth()
        {
            Application application = new Application();
            bool createdWtmlKey = false;

            try
            {
                createdWtmlKey = CreateWwtKeyIfNotExists();
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                Name namedRange = book.Names.GetNamedRange("GotoViewpointOnViewInWWTTestEarth");

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                Group sunGroup = new Group("Sun", GroupType.ReferenceFrame, null);

                LayerMap_Accessor selectedLayerMap = new LayerMap_Accessor(namedRange);
                selectedLayerMap.MapType = LayerMapType.LocalInWWT;
                selectedLayerMap.LayerDetails.Group = sunGroup;
                selectedLayerMap.LayerDetails.ID = "sunLayerID1";

                WorkflowController_Accessor.GotoViewpointOnViewInWWT(selectedLayerMap);
            }
            finally
            {
                application.Close();
                if (createdWtmlKey)
                {
                    // Delete the registry key if it is created by this test case.
                    Registry.ClassesRoot.DeleteSubKey(@".wtml", false);
                }
            }
        }

        /// <summary>
        /// A test for GotoViewpointOnViewInWWT
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GotoViewpointOnViewInWWTTestSky()
        {
            Application application = new Application();
            bool createdWtmlKey = false;

            try
            {
                createdWtmlKey = CreateWwtKeyIfNotExists();
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                Name namedRange = book.Names.GetNamedRange("GotoViewpointOnViewInWWTSky");

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                Group skyGroup = new Group("Sky", GroupType.ReferenceFrame, null);

                LayerMap_Accessor selectedLayerMap = new LayerMap_Accessor(namedRange);
                selectedLayerMap.MapType = LayerMapType.LocalInWWT;
                selectedLayerMap.LayerDetails.Group = skyGroup;
                selectedLayerMap.LayerDetails.ID = "sunLayerID1";

                WorkflowController_Accessor.GotoViewpointOnViewInWWT(selectedLayerMap);
            }
            finally
            {
                application.Close();
                if (createdWtmlKey)
                {
                    // Delete the registry key if it is created by this test case.
                    Registry.ClassesRoot.DeleteSubKey(@".wtml", false);
                }
            }
        }

        /// <summary>
        /// A test for GotoViewpointOnViewInWWT
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GotoViewpointOnViewInWWTTestEarthGeo()
        {
            Application application = new Application();
            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                Name namedRange = book.Names.GetNamedRange("GotoViewpointOnViewInWWTGeo");

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                Group sunGroup = new Group("Sun", GroupType.ReferenceFrame, null);

                LayerMap_Accessor selectedLayerMap = new LayerMap_Accessor(namedRange);
                selectedLayerMap.MapType = LayerMapType.LocalInWWT;
                selectedLayerMap.LayerDetails.Group = sunGroup;
                selectedLayerMap.LayerDetails.ID = "sunLayerID1";

                WorkflowController_Accessor.GotoViewpointOnViewInWWT(selectedLayerMap);
            }
            finally
            {
                application.Close();
            }
        }
        #endregion

        #region ValidateLocalInWWTLayerData

         /// <summary>
        /// A test for ValidateLocalInWWTLayerData
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void ValidateLocalInWWTLayerDataTest()
        {
            Application application = new Application();
            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);

                Name namedRange = book.Names.GetNamedRange("ColumnList");

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                LayerMap_Accessor localLayerMap = new LayerMap_Accessor(namedRange);
                localLayerMap.LayerDetails.Name = "Layer";
                localLayerMap.MapType = LayerMapType.LocalInWWT;

                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                target.currentWorkbookMap = new WorkbookMap_Accessor(book);
                target.currentWorkbookMap.SelectedLayerMap = localLayerMap;

                // Active sheet returned will be the book.Sheets[3] object. 
                book.Sheets[3].Activate();
                Range selectedRange = namedRange.RefersToRange;

                _Worksheet activeSheet = WorkflowController_Accessor.GetActiveWorksheet();

                bool expected = true;
                bool actual = target.ValidateLocalInWWTLayerData(activeSheet, selectedRange);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                application.Close();
            }
        }
        #endregion

        #region SetLayerDetailsViewModelProperties

         /// <summary>
        /// A test for SetLayerDetailsViewModelProperties
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void SetLayerDetailsViewModelPropertiesTest()
        {
            Application application = new Application();
            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);

                Name namedRange = book.Names.GetNamedRange("ColumnList");

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                LayerMap_Accessor localLayerMap = new LayerMap_Accessor(namedRange);
                localLayerMap.LayerDetails.Name = "Layer";
                localLayerMap.MapType = LayerMapType.LocalInWWT;

                LayerMap currentLayer = new LayerMap(namedRange);
                currentLayer.LayerDetails.Name = "Layer";
                currentLayer.MapType = LayerMapType.Local;

                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);

                target.currentWorkbookMap = new WorkbookMap_Accessor(book);
                target.layerDetailsViewModel = new LayerDetailsViewModel();
                target.layerDetailsViewModel.Currentlayer = currentLayer;
                target.currentWorkbookMap.SelectedLayerMap = localLayerMap;

                string expected = "Layer (not linked)";
                target.SetLayerDetailsViewModelProperties();

                Assert.AreEqual(expected, target.layerDetailsViewModel.SelectedLayerText);
                Assert.AreEqual(false, target.layerDetailsViewModel.IsViewInWWTEnabled);
                Assert.AreEqual(false, target.layerDetailsViewModel.IsCallOutVisible);
                Assert.AreEqual(true, target.layerDetailsViewModel.IsGetLayerDataEnabled);
                Assert.AreEqual(true, target.layerDetailsViewModel.IsReferenceGroupEnabled);
            }
            finally
            {
                application.Close();
            }
        }

        #region IsLocalLayer

        /// <summary>
        /// A test for IsLocalLayer
        /// </summary>
        [TestMethod()]
        public void IsLocalLayerTest()
        {
            bool actual = WorkflowController_Accessor.IsLocalLayer(LayerMapType.Local);
            Assert.AreEqual(true, actual);

            actual = WorkflowController_Accessor.IsLocalLayer(LayerMapType.LocalInWWT);
            Assert.AreEqual(true, actual);

            actual = WorkflowController_Accessor.IsLocalLayer(LayerMapType.WWT);
            Assert.AreEqual(false, actual);
        }

        #endregion

        /// <summary>
        /// A test for DeleteMapping
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void DeleteMappingTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                target.currentWorkbookMap = new WorkbookMap_Accessor(book);
                Worksheet worksheet = application.ActiveSheet;
                Range address = worksheet.get_Range("A1:G8");
                Name namedRange = book.CreateNamedRange("DeleteMappingRange", address);
                LayerMap_Accessor localLayer = new LayerMap_Accessor(namedRange);
                target.currentWorkbookMap.SelectedLayerMap = localLayer;
                target.DeleteMapping();
                Assert.IsNull(book.Names.GetNamedRange("DeleteMappingRange"));
                Assert.IsNull(target.currentWorkbookMap.SelectedLayerMap);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion

        #region ShowSelectedRangeTest

        /// <summary>
        /// A test for ShowSelectedRange
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void ShowSelectedRangeActiveSheetTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);

                target.currentWorkbookMap = new WorkbookMap_Accessor(book);
                Worksheet worksheet = application.ActiveSheet;
                Range inputRange = worksheet.get_Range("A1:G8");
                Name namedRange = book.CreateNamedRange("ShowSelectedRange", inputRange);
                LayerMap_Accessor localLayer = new LayerMap_Accessor(namedRange);
                target.currentWorkbookMap.SelectedLayerMap = localLayer;
                target.ShowSelectedRange();
                Range selectedRange = application.Selection as Range;
                Assert.AreEqual(selectedRange.Address, inputRange.Address);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for ShowSelectedRange
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void ShowSelectedRangeNonActiveSheetTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);

                target.currentWorkbookMap = new WorkbookMap_Accessor(book);
                Worksheet worksheet = application.Worksheets[1];
                Range inputRange = worksheet.get_Range("A1:G8");
                Name namedRange = book.CreateNamedRange("ShowSelectedRange", inputRange);
                LayerMap_Accessor localLayer = new LayerMap_Accessor(namedRange);
                target.currentWorkbookMap.SelectedLayerMap = localLayer;
                application.Worksheets[2].Activate();
                target.ShowSelectedRange();
                Range selectedRange = application.Selection as Range;
                Assert.AreEqual(selectedRange.Address, inputRange.Address);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion ShowSelectedRangeTest

        #region Workbook Events

        /// <summary>
        /// A test for OnWorkbookOpen
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnWorkbookOpenTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");
                target.OnWorkbookOpen(book);
                Assert.IsNotNull(target.currentWorkbookMap);
                Assert.AreEqual(target.currentWorkbookMap.Workbook, book);

                // CurrentWorkbookMap cannot be accessed directly through WorkflowController_Accessor object.
                WorkbookMap currentWorkbookMap = target.currentWorkbookMap.Target as WorkbookMap;

                Assert.AreEqual(currentWorkbookMap.AllLayerMaps.Count, 2);
                Assert.AreEqual(currentWorkbookMap.LocalLayerMaps.Count, 2);
                Assert.AreEqual(currentWorkbookMap.LocalInWWTLayerMaps.Count, 0);
                Assert.AreEqual(currentWorkbookMap.SerializableLayerMaps.Count, 2);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for OnSheetDeactivate
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnSheetDeactivateNoAffectedLayersTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");
                target.OnWorkbookOpen(book);
                target.OnSheetDeactivate(book.Worksheets[1]);

                // Affected layers is not accessible, but the code can be covered.
                Assert.AreEqual(target.mostRecentWorksheet, book.Worksheets[1]);
                Assert.AreEqual(target.mostRecentWorkbook, book);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for OnSheetDeactivate
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnSheetDeactivateWithAffectedLayersTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");
                target.OnWorkbookOpen(book);
                target.OnSheetDeactivate(book.Worksheets[11]);

                // Affected layers is not accessible, but the code can be covered.
                Assert.AreEqual(target.mostRecentWorksheet, book.Worksheets[11]);
                Assert.AreEqual(target.mostRecentWorkbook, book);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for OnNewWorkbook
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnNewWorkbookTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.Workbooks.Add();
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");
                target.OnNewWorkbook(book);

                Assert.IsNotNull(target.currentWorkbookMap);
                Assert.AreEqual(target.currentWorkbookMap.Workbook, book);

                // CurrentWorkbookMap cannot be accessed directly through WorkflowController_Accessor object.
                WorkbookMap currentWorkbookMap = target.currentWorkbookMap.Target as WorkbookMap;

                Assert.AreEqual(currentWorkbookMap.AllLayerMaps.Count, 0);
                Assert.AreEqual(currentWorkbookMap.LocalLayerMaps.Count, 0);
                Assert.AreEqual(currentWorkbookMap.LocalInWWTLayerMaps.Count, 0);
                Assert.AreEqual(currentWorkbookMap.SerializableLayerMaps.Count, 0);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for SheetChange
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnSheetChangeTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");
                target.OnWorkbookOpen(book);

                Name targetRangeName = book.Names.GetNamedRange("SetAllPropertiesRange");
                targetRangeName.RefersToRange.Cells[1, 1].Value = "Modified";
                target.OnSheetChange(targetRangeName.RefersToRange.Worksheet, targetRangeName.RefersToRange);

                // CurrentWorkbookMap cannot be accessed directly through WorkflowController_Accessor object.
                WorkbookMap currentWorkbookMap = target.currentWorkbookMap.Target as WorkbookMap;
                Assert.IsNotNull(currentWorkbookMap);
                List<LayerMap> allLayerMaps = currentWorkbookMap.AllLayerMaps as List<LayerMap>;
                Assert.IsNotNull(allLayerMaps);
                var layerMapExpected = allLayerMaps.Find(layerMap => layerMap.RangeDisplayName == "TestProperties_1");
                Assert.AreEqual(layerMapExpected.HeaderRowData[0], "Modified");
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for OnWorkbookActivate
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnWorkbookActivateTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");
                target.OnWorkbookOpen(book);
                target.OnWorkbookActivate(book);
                Assert.IsNotNull(target.currentWorkbookMap);
                Assert.AreEqual("WorkbookTestData.xlsx", target.currentWorkbookMap.Workbook.Name);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion Workbook Events

        #region GetSelectedLayerMapTest

        /// <summary>
        /// A test for GetSelectedLayerMap
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetSelectedLayerMapWWTTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");
                target.OnNewWorkbook(book);

                Layer layer = new Layer();
                LayerMap_Accessor localLayer = new LayerMap_Accessor(layer);
                localLayer.MapType = LayerMapType.WWT;
                localLayer.LayerDetails.ID = "2cf4374f-e1ce-47a9-b08c-31079765ddcf";

                // CurrentWorkbookMap cannot be accessed directly through WorkflowController_Accessor object.
                WorkbookMap currentWorkbookMap = target.currentWorkbookMap.Target as WorkbookMap;
                currentWorkbookMap.AllLayerMaps[0].LayerDetails.ID = "2cf4374f-e1ce-47a9-b08c-31079765ddcf";
                currentWorkbookMap.AllLayerMaps[0].MapType = LayerMapType.WWT;
                LayerMap_Accessor layerMapAccessor = target.GetSelectedLayerMap(localLayer);

                // Make sure LayerMap is returned and the expected LayerMap ID is returned.
                Assert.IsNotNull(layerMapAccessor);
                Assert.AreEqual(layerMapAccessor.LayerDetails.ID, "2cf4374f-e1ce-47a9-b08c-31079765ddcf");
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for GetSelectedLayerMap
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetSelectedLayerMapLocalTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");
                target.OnNewWorkbook(book);

                Name name = book.Names.GetNamedRange("ColumnList");
                LayerMap localLayer = new LayerMap(name);
                localLayer.MapType = LayerMapType.Local;
                localLayer.LayerDetails.ID = "2cf4374f-e1ce-47a9-b08c-31079765ddcf";

                // CurrentWorkbookMap cannot be accessed directly through WorkflowController_Accessor object.
                WorkbookMap currentWorkbookMap = target.currentWorkbookMap.Target as WorkbookMap;
                currentWorkbookMap.AllLayerMaps.Add(localLayer);

                LayerMap_Accessor layerMapAccessor = new LayerMap_Accessor(name);
                localLayer.MapType = LayerMapType.Local;
                localLayer.LayerDetails.ID = "2cf4374f-e1ce-47a9-b08c-31079765ddcf";
                layerMapAccessor = target.GetSelectedLayerMap(layerMapAccessor);

                // Make sure LayerMap is returned and the expected LayerMap ID is returned.
                Assert.IsNotNull(layerMapAccessor);
                Assert.AreEqual(layerMapAccessor.LayerDetails.ID, "2cf4374f-e1ce-47a9-b08c-31079765ddcf");
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for GetSelectedLayerMap
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetSelectedLayerMapLocalInWWTTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");
                target.OnNewWorkbook(book);

                Layer layer = new Layer();
                LayerMap_Accessor localLayer = new LayerMap_Accessor(layer);
                localLayer.MapType = LayerMapType.LocalInWWT;
                localLayer.LayerDetails.ID = "2cf4374f-e1ce-47a9-b08c-31079765ddcf";

                // CurrentWorkbookMap cannot be accessed directly through WorkflowController_Accessor object.
                WorkbookMap currentWorkbookMap = target.currentWorkbookMap.Target as WorkbookMap;
                currentWorkbookMap.AllLayerMaps[0].LayerDetails.ID = "2cf4374f-e1ce-47a9-b08c-31079765ddcf";
                currentWorkbookMap.AllLayerMaps[0].MapType = LayerMapType.LocalInWWT;
                LayerMap_Accessor layerMapAccessor = target.GetSelectedLayerMap(localLayer);

                // Make sure LayerMap is returned and the expected LayerMap ID is returned.
                Assert.IsNotNull(layerMapAccessor);
                Assert.AreEqual(layerMapAccessor.LayerDetails.ID, "2cf4374f-e1ce-47a9-b08c-31079765ddcf");
            }
            finally
            {
                application.Close();
            }
        }

        #endregion GetSelectedLayerMapTest

        #region SetFormatForDateColumns

        /// <summary>
        /// A test for SetFormatForDateColumns
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void SetFormatForDateColumnsTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                Name name = book.Names.GetNamedRange("SetFormatForDateColumns_1");

                foreach (Range cell in name.RefersToRange.Range["D2:E6"])
                {
                    Assert.AreEqual(cell.NumberFormat, "m/d/yyyy");
                }

                target.OnWorkbookOpen(book);

                target.SetFormatForDateColumns(book.ActiveSheet);

                foreach (Range cell in name.RefersToRange.Range["D2:E6"])
                {
                    Assert.AreEqual(cell.NumberFormat, "m/d/yyyy h:mm");
                }
            }
            finally
            {
                application.Close();
            }
        }

        #endregion SetFormatForDateColumns

        #region Events

        #region OnCustomTaskPaneChangedState

        /// <summary>
        /// A test for OnCustomTaskPaneChangedState
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnCustomTaskPaneChangedStateTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);

                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                target.currentWorkbookMap = new WorkbookMap_Accessor(book);

                // Get the named range stored in the test data excel file.
                Name name = book.Names.GetNamedRange("ColumnList");
                Group parentGroup = new Group("Sun", GroupType.ReferenceFrame, null);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                LayerMap_Accessor localLayerMap = new LayerMap_Accessor(name);
                localLayerMap.LayerDetails.Name = "Layer";
                localLayerMap.LayerDetails.Group = parentGroup;
                localLayerMap.MapType = LayerMapType.LocalInWWT;

                target.currentWorkbookMap.SelectedLayerMap = localLayerMap;
                target.layerDetailsViewModel = new LayerDetailsViewModel();

                target.layerDetailsViewModel.ColumnsView = new System.Collections.ObjectModel.ObservableCollection<ColumnViewModel>();
                target.layerDetailsViewModel.ColumnsView.Add(new ColumnViewModel()
                {
                    ExcelHeaderColumn = "Header",
                    SelectedWWTColumn = new Column(ColumnType.Alt, "ss", new System.Collections.ObjectModel.Collection<string>() { "s" })
                });

                LayerMap currentLayer = new LayerMap(name);
                currentLayer.LayerDetails.Name = "Layer";
                currentLayer.LayerDetails.Group = parentGroup;
                currentLayer.MapType = LayerMapType.LocalInWWT; 

                target.layerDetailsViewModel.Currentlayer = currentLayer;
                target.layerDetailsViewModel.SelectedLayerName = "Layer1";
                target.layerDetailsViewModel.SelectedGroup = parentGroup;
                target.layerDetailsViewModel.IsDistanceVisible = true;
                target.layerDetailsViewModel.SelectedDistanceUnit = new KeyValuePair<AltUnit, string>(AltUnit.AstronomicalUnits, "Alt");
                target.layerDetailsViewModel.IsRAUnitVisible = true;
                target.layerDetailsViewModel.SelectedFadeType = new KeyValuePair<FadeType, string>(FadeType.Both, "f");
                target.layerDetailsViewModel.SelectedScaleType = new KeyValuePair<ScaleType, string>(ScaleType.Constant, "Constant");
                target.layerDetailsViewModel.SelectedScaleRelative = new KeyValuePair<ScaleRelativeType, string>(ScaleRelativeType.Screen, "Screen");
                target.layerDetailsViewModel.ScaleFactor.SelectedSliderValue = 10;
                target.layerDetailsViewModel.LayerOpacity.SelectedSliderValue = 10;
                target.layerDetailsViewModel.SelectedSize = new KeyValuePair<int, string>(1, "2");
                target.layerDetailsViewModel.SelectedHoverText = new KeyValuePair<int, string>(1, "2");

                object sender = null;
                EventArgs e = null;
                target.OnCustomTaskPaneChangedState(sender, e);

                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.LayerDetails.Group, parentGroup);
                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.name, localLayerMap.name);
                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.LayerDetails.AltUnit, target.layerDetailsViewModel.SelectedDistanceUnit.Key);
                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.LayerDetails.FadeType, target.layerDetailsViewModel.SelectedFadeType.Key);
                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.LayerDetails.PointScaleType, target.layerDetailsViewModel.SelectedScaleType.Key);
                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.LayerDetails.MarkerScale, target.layerDetailsViewModel.SelectedScaleRelative.Key);
                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.LayerDetails.ScaleFactor, 0.125);
                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.LayerDetails.Opacity, target.layerDetailsViewModel.LayerOpacity.SelectedSliderValue / 100);
                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.LayerDetails.SizeColumn, target.layerDetailsViewModel.SelectedSize.Key);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion

        #region OnReferenceSelectionChangedest

        /// <summary>
        /// A test for OnReferenceSelectionChanged
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnReferenceSelectionChangedTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                target.currentWorkbookMap = new WorkbookMap_Accessor(book);

                // Get the named range stored in the test data excel file.
                Name name = book.Names.GetNamedRange("ColumnList");
                Group parentGroup = new Group("Sun", GroupType.ReferenceFrame, null);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                string layerId = "2cf4374f-e1ce-47a9-b08c-31079765ddcf";
                LayerMap_Accessor localLayerMap = new LayerMap_Accessor(name);
                localLayerMap.LayerDetails.Name = "Layer1";
                localLayerMap.LayerDetails.ID = layerId;
                localLayerMap.LayerDetails.Group = parentGroup;
                localLayerMap.MapType = LayerMapType.LocalInWWT;

                LayerMap currentlayer = new LayerMap(name);
                currentlayer.LayerDetails.Name = "Layer2";
                currentlayer.LayerDetails.ID = layerId;
                currentlayer.LayerDetails.Group = parentGroup;
                currentlayer.MapType = LayerMapType.LocalInWWT;

                currentlayer.HeaderRowData = new Collection<string>(); 
                currentlayer.HeaderRowData.Add("LAT");
                currentlayer.HeaderRowData.Add("Long");
                currentlayer.HeaderRowData.Add("Depth");

                target.layerDetailsViewModel = new LayerDetailsViewModel(); 
                target.layerDetailsViewModel.Currentlayer = currentlayer;
                target.currentWorkbookMap.SelectedLayerMap = localLayerMap;
                EventArgs e = null;
                target.OnReferenceSelectionChanged(null, e);

                Collection<ColumnType> mappedType = new Collection<ColumnType>();
                mappedType.Add(ColumnType.Lat);
                mappedType.Add(ColumnType.Long);
                mappedType.Add(ColumnType.Depth);

                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.LayerDetails.Name, target.layerDetailsViewModel.Currentlayer.LayerDetails.Name);
                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.MappedColumnType.Count, target.layerDetailsViewModel.Currentlayer.MappedColumnType.Count);

                int index = 0;
                foreach (ColumnType colType in mappedType)
                {
                    Assert.AreEqual(target.layerDetailsViewModel.Currentlayer.MappedColumnType[index], colType);
                    index++;
                }
            }
            finally
            {
                application.Close();
            }
        }
        #endregion

        #region OnViewInWWTClicked

        /// <summary>
        /// A test for OnViewInWWTClicked
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnViewInWWTClickedTest()
        {
            Application application = new Application();
            bool createdWtmlKey = false;

            try
            {
                // .wtml registry key is needed to be there for this test case to pass. Key will be created and deleted, if it doesn't exist.
                createdWtmlKey = CreateWwtKeyIfNotExists();
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);

                target.currentWorkbookMap = new WorkbookMap_Accessor(book);

                // Get the named range stored in the test data excel file.
                Name name = book.Names.GetNamedRange("ColumnList");
                Group parentGroup = new Group("Sun", GroupType.ReferenceFrame, null);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                string layerId = "2cf4374f-e1ce-47a9-b08c-31079765ddcf";
                LayerMap_Accessor localLayerMap = new LayerMap_Accessor(name);
                localLayerMap.LayerDetails.Name = "Layer1";
                localLayerMap.LayerDetails.ID = layerId;
                localLayerMap.LayerDetails.Group = parentGroup;
                localLayerMap.MapType = LayerMapType.LocalInWWT;

                parentGroup.LayerIDs.Add(layerId);

                LayerMap currentlayer = new LayerMap(name);
                currentlayer.LayerDetails.Name = "Layer2";
                currentlayer.LayerDetails.ID = layerId;
                currentlayer.LayerDetails.Group = parentGroup;
                currentlayer.MapType = LayerMapType.LocalInWWT;

                currentlayer.HeaderRowData = new Collection<string>();
                currentlayer.HeaderRowData.Add("LAT");
                currentlayer.HeaderRowData.Add("Long");
                currentlayer.HeaderRowData.Add("Depth");

                target.layerDetailsViewModel = new LayerDetailsViewModel();
                target.layerDetailsViewModel.Currentlayer = currentlayer;
                target.layerDetailsViewModel.SelectedLayerName = "Layer1";

                Collection<ColumnType> mappedType = new Collection<ColumnType>();
                mappedType.Add(ColumnType.Lat);
                mappedType.Add(ColumnType.Long);
                mappedType.Add(ColumnType.Depth);

                target.currentWorkbookMap.SelectedLayerMap = localLayerMap;
                target.currentWorkbookMap.SelectedLayerMap.MappedColumnType = mappedType;

                EventArgs e = null;
                target.OnViewInWWTClicked(null, e);
                string expected = "Layer1 (linked)";

                Assert.AreEqual(layerId, target.currentWorkbookMap.SelectedLayerMap.LayerDetails.ID);
                Assert.AreEqual(expected, target.layerDetailsViewModel.SelectedLayerText);
                Assert.AreEqual(false, target.layerDetailsViewModel.IsViewInWWTEnabled);
                Assert.AreEqual(false, target.layerDetailsViewModel.IsCallOutVisible);
                Assert.AreEqual(true, target.layerDetailsViewModel.IsGetLayerDataEnabled);
                Assert.AreEqual(false, target.layerDetailsViewModel.IsReferenceGroupEnabled);
                Assert.AreEqual("Refresh", target.layerDetailsViewModel.LayerDataDisplayName);

                localLayerMap.LayerDetails.ID = string.Empty;
                localLayerMap.MapType = LayerMapType.Local;
                target.OnViewInWWTClicked(null, e);

                Assert.AreEqual(layerId, target.currentWorkbookMap.SelectedLayerMap.LayerDetails.ID);
                Assert.AreEqual(expected, target.layerDetailsViewModel.SelectedLayerText);
                Assert.AreEqual(false, target.layerDetailsViewModel.IsViewInWWTEnabled);
                Assert.AreEqual(false, target.layerDetailsViewModel.IsCallOutVisible);
                Assert.AreEqual(true, target.layerDetailsViewModel.IsGetLayerDataEnabled);
                Assert.AreEqual(false, target.layerDetailsViewModel.IsReferenceGroupEnabled);
                Assert.AreEqual("Refresh", target.layerDetailsViewModel.LayerDataDisplayName);
            }
            finally
            {
                application.Close();
                if (createdWtmlKey)
                {
                    // Delete the registry key if it is created by this test case.
                    Registry.ClassesRoot.DeleteSubKey(@".wtml", false);
                }
            }
        }
        
        #endregion

        #region OnRefreshDropDownClickedEvent

        /// <summary>
        /// A test for OnRefreshDropDownClickedEvent
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnRefreshDropDownClickedEventTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                target.currentWorkbookMap = new WorkbookMap_Accessor(book);

                // Get the named range stored in the test data excel file.
                Name name = book.Names.GetNamedRange("ColumnList");
                Group parentGroup = new Group("Sun", GroupType.ReferenceFrame, null);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                string layerId = "2cf4374f-e1ce-47a9-b08c-31079765ddcf";
                LayerMap_Accessor localLayerMap = new LayerMap_Accessor(name);
                localLayerMap.LayerDetails.Name = "Layer1";
                localLayerMap.LayerDetails.ID = layerId;
                localLayerMap.LayerDetails.Group = parentGroup;
                localLayerMap.MapType = LayerMapType.LocalInWWT;

                target.currentWorkbookMap.SelectedLayerMap = localLayerMap;

                EventArgs e = null;
                target.OnRefreshDropDownClickedEvent(null, e);

                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.LayerDetails.Name, localLayerMap.LayerDetails.Name);
                Assert.AreEqual(target.layerDetailsViewModel.Currentlayer.LayerDetails.Name, target.currentWorkbookMap.SelectedLayerMap.LayerDetails.Name);
                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.LayerDetails.Name, target.layerDetailsViewModel.SelectedLayerName);
                Assert.AreEqual(target.layerDetailsViewModel.SelectedGroupText, target.currentWorkbookMap.SelectedLayerMap.LayerDetails.Group.Name);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for OnRefreshGroupDropDownClickedEvent
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnRefreshGroupDropDownClickedEventTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                target.currentWorkbookMap = new WorkbookMap_Accessor(book);

                // Get the named range stored in the test data excel file.
                Name name = book.Names.GetNamedRange("ColumnList");
                Group parentGroup = new Group("Sun", GroupType.ReferenceFrame, null);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                string layerId = "2cf4374f-e1ce-47a9-b08c-31079765ddcf";
                LayerMap_Accessor localLayerMap = new LayerMap_Accessor(name);
                localLayerMap.LayerDetails.Name = "Layer1";
                localLayerMap.LayerDetails.ID = layerId;
                localLayerMap.LayerDetails.Group = parentGroup;
                localLayerMap.MapType = LayerMapType.LocalInWWT;
                target.currentWorkbookMap.SelectedLayerMap = localLayerMap;

                EventArgs e = null;
                target.OnRefreshGroupDropDownClickedEvent(null, e);

                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.LayerDetails.Name, localLayerMap.LayerDetails.Name);
                Assert.AreEqual(target.layerDetailsViewModel.Currentlayer.LayerDetails.Name, target.currentWorkbookMap.SelectedLayerMap.LayerDetails.Name);
                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.LayerDetails.Name, target.layerDetailsViewModel.SelectedLayerName);
                Assert.AreEqual(target.layerDetailsViewModel.SelectedGroupText, target.currentWorkbookMap.SelectedLayerMap.LayerDetails.Group.Name);
            }
            finally
            {
                application.Close();
            }
        }
        #endregion

        #region OnGetLayerDataClickedEvent

        /// <summary>
        /// A test for OnGetLayerDataClickedEvent invalid layer
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnGetLayerDataClickedEventInvalidLayerTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);

                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                target.currentWorkbookMap = new WorkbookMap_Accessor(book);

                // Get the named range stored in the test data excel file.
                Name name = book.Names.GetNamedRange("ColumnList");
                Group parentGroup = new Group("Sun", GroupType.ReferenceFrame, null);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTInvalidMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                string layerId = "2cf4374f-e1ce-47a9-b08c-31079765ddcf";
                LayerMap_Accessor localLayerMap = new LayerMap_Accessor(name);
                localLayerMap.LayerDetails.Name = "Layer1";
                localLayerMap.LayerDetails.ID = layerId;
                localLayerMap.LayerDetails.Group = parentGroup;
                localLayerMap.MapType = LayerMapType.LocalInWWT;

                target.currentWorkbookMap.SelectedLayerMap = localLayerMap;

                EventArgs e = null;
                target.OnGetLayerDataClickedEvent(null, e);

                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.LayerDetails.Name, localLayerMap.LayerDetails.Name);
                Assert.AreEqual(target.layerDetailsViewModel.Currentlayer.LayerDetails.Name, target.currentWorkbookMap.SelectedLayerMap.LayerDetails.Name);
                Assert.AreEqual(target.currentWorkbookMap.SelectedLayerMap.LayerDetails.Name, target.layerDetailsViewModel.SelectedLayerName);
                Assert.AreEqual(target.layerDetailsViewModel.SelectedGroupText, target.currentWorkbookMap.SelectedLayerMap.LayerDetails.Group.Name);
            }
            finally
            {
                application.Close();
            }
        }
        #endregion

        #region OnTargetMachineChanged

          /// <summary>
        /// A test for OnTargetMachineChanged
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnTargetMachineChangedTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("TestData.xlsx", false);

                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                target.currentWorkbookMap = new WorkbookMap_Accessor(book);

                // Get the named range stored in the test data excel file.
                Name name = book.Names.GetNamedRange("ColumnList");
                Group parentGroup = new Group("Sun", GroupType.ReferenceFrame, null);

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTInvalidMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                string layerId = "2cf4374f-e1ce-47a9-b08c-31079765ddcf";
                LayerMap_Accessor wwtLayerMap = new LayerMap_Accessor(name);
                wwtLayerMap.LayerDetails.Name = "Layer1";
                wwtLayerMap.LayerDetails.ID = layerId;
                wwtLayerMap.LayerDetails.Group = parentGroup;
                wwtLayerMap.MapType = LayerMapType.WWT;

                target.currentWorkbookMap.SelectedLayerMap = wwtLayerMap;

                EventArgs e = null;
                target.OnTargetMachineChanged(null, e);

                Assert.AreEqual(null, target.currentWorkbookMap.SelectedLayerMap);
                Assert.AreEqual(null, target.layerDetailsViewModel.Currentlayer);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion

        #region OnUpdateLayerClickedEvent

        /// <summary>
        /// A test for OnTargetMachineChanged
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnUpdateLayerClickedEventTest()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");
                target.OnWorkbookOpen(book);

                Name targetRangeName = book.Names.GetNamedRange("UpdateLayerRangeInitial");
                
                WorkbookMap_Accessor workbookMap = Microsoft.Research.Wwt.Excel.Addin.WorkbookExtensions_Accessor.GetWorkbookMap(book);
                LayerMap_Accessor layerMap = new LayerMap_Accessor(targetRangeName);

                workbookMap.SelectedLayerMap = layerMap;
                target.currentWorkbookMap = workbookMap;

                Range updateRangeFinal = targetRangeName.RefersToRange.Worksheet.get_Range("A1:F15");
                updateRangeFinal.Select();
                target.OnUpdateLayerClickedEvent(this, new EventArgs());

                Assert.AreEqual("$A$1:$F$15", target.currentWorkbookMap.SelectedLayerMap.RangeName.RefersToRange.Address);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion

        #region OnVizualiseSelectionClicked

        /// <summary>
        /// A test for OnTargetMachineChanged
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void OnVisualizeSelectionClicked()
        {
            Application application = new Application();

            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                WorkflowController_Accessor target = GetWorkflowControllerAccessor(application);
                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");
                target.OnWorkbookOpen(book);

                Name targetRangeName = book.Names.GetNamedRange("UpdateLayerRangeInitial");
                Range updateRangeFinal = targetRangeName.RefersToRange.Worksheet.get_Range("A1:F15");
               
                updateRangeFinal.Select();
                target.OnVisualizeSelectionClicked(this, new EventArgs());
                
                // CurrentWorkbookMap cannot be accessed directly through WorkflowController_Accessor object.
                WorkbookMap currentWorkbookMap = target.currentWorkbookMap.Target as WorkbookMap;
                Assert.IsNotNull(currentWorkbookMap);
                List<LayerMap> allLayerMaps = currentWorkbookMap.AllLayerMaps as List<LayerMap>;
                Assert.IsNotNull(allLayerMaps);
                Assert.AreEqual(3, allLayerMaps.Count);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion

        #endregion

        #region Private Methods

        /// <summary>
        /// Creates the WWT registry key "HKEY_CLASSES_ROOT\.wtml" so that the test cases will treat as WWT installed on the machine.
        /// </summary>
        /// <returns>True if key is created. False, if it already exists.</returns>
        private static bool CreateWwtKeyIfNotExists()
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

            return createdWtmlKey;
        }

        /// <summary>
        /// Gets an instance of WorkflowController_Accessor. Also, sets the ExcelApplication object of ThisAddIn_Accessor.
        /// </summary>
        /// <param name="excelApplication">Application object</param>
        /// <returns>Instance of WorkflowController_Accessor</returns>
        private static WorkflowController_Accessor GetWorkflowControllerAccessor(Application excelApplication)
        {
            // Where ever we are creating object for WorkflowController_Accessor, we need to set the ThisAddIn_Accessor.ExcelApplication to 
            // wither null or actual application object.
            ThisAddIn_Accessor.ExcelApplication = excelApplication;

            WorkflowController_Accessor target = new WorkflowController_Accessor();
            return target;
        }

        #endregion Private Methods
    }
}