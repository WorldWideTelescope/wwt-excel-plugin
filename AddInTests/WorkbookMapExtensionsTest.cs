//-----------------------------------------------------------------------
// <copyright file="WorkbookMapExtensionsTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Collections.Generic;
using System.Collections.ObjectModel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Research.Wwt.Excel.Addin;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{
    /// <summary>
    /// This is a test class for WorkbookMapExtensionsTest and is intended
    /// to contain all WorkbookMapExtensionsTest Unit Tests
    /// </summary>
    [TestClass()]
    public class WorkbookMapExtensionsTest
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

        #region CleanUpWWTLayersTest

        /// <summary>
        /// A test for CleanUpWWTLayers
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void CleanUpWWTLayersTest()
        {
            Application application = new Application();
            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                Name emptyRange = book.Names.GetNamedRange("GetSelectedLayerWorksheet");

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                Group marsGroup = new Group("Mars", GroupType.ReferenceFrame, null);
                Group sunGroup = new Group("Sun", GroupType.ReferenceFrame, null);
                Group earthGroup = new Group("Earth", GroupType.ReferenceFrame, sunGroup);
                earthGroup.LayerIDs.Add("8e5cbbc4-9eb2-47e5-b3d7-7ca9babda477");
                sunGroup.Children.Add(earthGroup);

                LayerMap earthLayer = new LayerMap(new Layer());
                earthLayer.MapType = LayerMapType.WWT;
                earthLayer.LayerDetails.Group = earthGroup;
                earthLayer.LayerDetails.ID = "ValidID1";

                LayerMap marsLayer = new LayerMap(emptyRange);
                marsLayer.MapType = LayerMapType.LocalInWWT;
                marsLayer.LayerDetails.Group = marsGroup;
                marsLayer.LayerDetails.ID = "MarsValidID1";

                LayerMap sunLayer = new LayerMap(emptyRange);
                sunLayer.MapType = LayerMapType.LocalInWWT;
                sunLayer.LayerDetails.Group = sunGroup;
                sunLayer.LayerDetails.ID = "sunLayerID1";

                WorkbookMap workbookMap = new WorkbookMap(null);
                workbookMap.AllLayerMaps.Add(earthLayer);
                workbookMap.AllLayerMaps.Add(marsLayer);
                workbookMap.AllLayerMaps.Add(sunLayer);

                workbookMap.SelectedLayerMap = earthLayer;

                WorkbookMapExtensions.CleanUpWWTLayers(workbookMap);
                Assert.IsNull(workbookMap.SelectedLayerMap);
                Assert.IsTrue(workbookMap.AllLayerMaps.Count == 2);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion

        #region AddWWTLayersTest

        /// <summary>
        /// A test for AddWWTLayers
        /// </summary>
        [TestMethod()]
        public void AddWWTLayersTest()
        {
            Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
            Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

            Group sunGroup = new Group("Sun", GroupType.ReferenceFrame, null);
            Group earthGroup = new Group("Earth", GroupType.ReferenceFrame, sunGroup);
            earthGroup.LayerIDs.Add("8e5cbbc4-9eb2-47e5-b3d7-7ca9babda477");
            sunGroup.Children.Add(earthGroup);

            Workbook workbook = null;
            WorkbookMap_Accessor workbookMap = new WorkbookMap_Accessor(workbook);

            WorkbookMapExtensions_Accessor.AddWWTLayers(workbookMap, sunGroup);
        }

        #endregion

        #region UpdateGroupStatusTest
        /// <summary>
        /// A test for UpdateGroupStatus
        /// </summary>
        [TestMethod()]
        public void UpdateGroupStatusTest()
        {
            Group sunGroup = new Group("Sun", GroupType.ReferenceFrame, null);
            Group skyGroup = new Group("Sky", GroupType.ReferenceFrame, null);
            Group earthGroup = new Group("Earth", GroupType.ReferenceFrame, sunGroup);
            earthGroup.LayerIDs.Add("8e5cbbc4-9eb2-47e5-b3d7-7ca9babda477");
            sunGroup.Children.Add(earthGroup);

            Collection<Group> wwtGroups = new Collection<Group>();
            wwtGroups.Add(skyGroup);

            WorkbookMapExtensions_Accessor.UpdateGroupStatus(earthGroup, wwtGroups);
        }
        #endregion

        #region SearchGroupTest

        /// <summary>
        /// A test for SearchGroup
        /// </summary>
        [TestMethod()]
        public void SearchGroupTestWrongName()
        {
            Group sunGroup = new Group("Sun", GroupType.ReferenceFrame, null);
            Group skyGroup = new Group("Sky", GroupType.ReferenceFrame, null);
            Group earthGroup = new Group("Earth", GroupType.ReferenceFrame, sunGroup);
            earthGroup.LayerIDs.Add("8e5cbbc4-9eb2-47e5-b3d7-7ca9babda477");
            sunGroup.Children.Add(earthGroup);

            Collection<Group> wwtGroups = new Collection<Group>();
            wwtGroups.Add(sunGroup);
            wwtGroups.Add(skyGroup);

            Group expected = null;
            Group actual;
            actual = WorkbookMapExtensions_Accessor.SearchGroup("WrongName", string.Empty, wwtGroups);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for SearchGroup
        /// </summary>
        [TestMethod()]
        public void SearchGroupTestValidName()
        {
            string groupName = "Earth";
            string path = "/Sun/Earth";
            Group sunGroup = new Group("Sun", GroupType.ReferenceFrame, null);
            Group skyGroup = new Group("Sky", GroupType.ReferenceFrame, null);
            Group earthGroup = new Group("Earth", GroupType.ReferenceFrame, sunGroup);
            earthGroup.LayerIDs.Add("8e5cbbc4-9eb2-47e5-b3d7-7ca9babda477");
            sunGroup.Children.Add(earthGroup);

            Collection<Group> wwtGroups = new Collection<Group>();
            wwtGroups.Add(sunGroup);
            wwtGroups.Add(skyGroup);

            Group actual = WorkbookMapExtensions_Accessor.SearchGroup(groupName, path, wwtGroups);
            Assert.AreEqual(groupName, actual.Name);
            Assert.AreEqual(path, actual.Path);
        }

        #endregion

        #region ExistsTest

        /// <summary>
        /// A test for Exists
        /// </summary>
        [TestMethod()]
        public void ExistsTestEmptyLayerId()
        {
            WorkbookMap workbookMap = new WorkbookMap(null);
            string layerID = string.Empty;
            bool expected = false;
            bool actual;
            actual = WorkbookMapExtensions.Exists(workbookMap, layerID);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for Exists
        /// </summary>
        [TestMethod()]
        public void ExistsTestEmptyWorkbookMap()
        {
            string layerID = string.Empty;
            bool expected = false;
            bool actual;
            actual = WorkbookMapExtensions.Exists(null, layerID);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for Exists
        /// </summary>
        [TestMethod()]
        public void ExistsTestValid()
        {
            LayerMap layerMap = new LayerMap(new Layer());
            layerMap.LayerDetails.ID = "ValidID";

            WorkbookMap workbookMap = new WorkbookMap(null);
            workbookMap.AllLayerMaps.Add(layerMap);

            string layerID = "ValidID";
            bool expected = true;
            bool actual;
            actual = WorkbookMapExtensions.Exists(workbookMap, layerID);
            Assert.AreEqual(expected, actual);
        }

        #endregion

        #region LoadWWTLayersTest

        /// <summary>
        /// A test for LoadWWTLayers
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void LoadWWTLayersTest()
        {
            Application application = new Application();
            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                Name emptyRange = book.Names.GetNamedRange("GetSelectedLayerWorksheet");

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                Group marsGroup = new Group("Mars", GroupType.ReferenceFrame, null);
                Group sunGroup = new Group("Sun", GroupType.ReferenceFrame, null);
                Group earthGroup = new Group("Earth", GroupType.ReferenceFrame, sunGroup);
                earthGroup.LayerIDs.Add("8e5cbbc4-9eb2-47e5-b3d7-7ca9babda477");
                sunGroup.Children.Add(earthGroup);

                LayerMap earthLayer = new LayerMap(new Layer());
                earthLayer.MapType = LayerMapType.WWT;
                earthLayer.LayerDetails.Group = earthGroup;
                earthLayer.LayerDetails.ID = "ValidID1";

                LayerMap marsLayer = new LayerMap(emptyRange);
                marsLayer.MapType = LayerMapType.LocalInWWT;
                marsLayer.LayerDetails.Group = marsGroup;
                marsLayer.LayerDetails.ID = "MarsValidID1";

                LayerMap sunLayer = new LayerMap(emptyRange);
                sunLayer.MapType = LayerMapType.LocalInWWT;
                sunLayer.LayerDetails.Group = sunGroup;
                sunLayer.LayerDetails.ID = "sunLayerID1";

                WorkbookMap workbookMap = new WorkbookMap(null);
                workbookMap.AllLayerMaps.Add(earthLayer);
                workbookMap.AllLayerMaps.Add(marsLayer);
                workbookMap.AllLayerMaps.Add(sunLayer);

                workbookMap.SelectedLayerMap = earthLayer;

                WorkbookMapExtensions.LoadWWTLayers(workbookMap);
                Assert.IsNull(workbookMap.SelectedLayerMap);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion

        #region RefreshLayersTest

        /// <summary>
        /// A test for RefreshLayers
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void RefreshLayersTest()
        {
            Application application = new Application();
            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                Name emptyRange = book.Names.GetNamedRange("GetSelectedLayerWorksheet");

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                Group marsGroup = new Group("Mars", GroupType.ReferenceFrame, null);
                Group sunGroup = new Group("Sun", GroupType.ReferenceFrame, null);
                Group earthGroup = new Group("Earth", GroupType.ReferenceFrame, sunGroup);
                earthGroup.LayerIDs.Add("8e5cbbc4-9eb2-47e5-b3d7-7ca9babda477");
                sunGroup.Children.Add(earthGroup);

                LayerMap earthLayer = new LayerMap(new Layer());
                earthLayer.MapType = LayerMapType.WWT;
                earthLayer.LayerDetails.Group = earthGroup;
                earthLayer.LayerDetails.ID = "ValidID1";

                LayerMap marsLayer = new LayerMap(emptyRange);
                marsLayer.MapType = LayerMapType.LocalInWWT;
                marsLayer.LayerDetails.Group = marsGroup;
                marsLayer.LayerDetails.ID = "MarsValidID1";

                LayerMap sunLayer = new LayerMap(emptyRange);
                sunLayer.MapType = LayerMapType.LocalInWWT;
                sunLayer.LayerDetails.Group = sunGroup;
                sunLayer.LayerDetails.ID = "sunLayerID1";

                WorkbookMap workbookMap = new WorkbookMap(null);
                workbookMap.AllLayerMaps.Add(earthLayer);
                workbookMap.AllLayerMaps.Add(marsLayer);
                workbookMap.AllLayerMaps.Add(sunLayer);

                workbookMap.SelectedLayerMap = earthLayer;

                WorkbookMapExtensions.RefreshLayers(workbookMap);
                Assert.IsNull(workbookMap.SelectedLayerMap);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion

        #region GetNamedRangesForInSyncLayersTest

        /// <summary>
        /// A test for GetNamedRangesForInSyncLayers
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void GetNamedRangesForInSyncLayersTest()
        {
            Application application = new Application();
            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                Name emptyRange = book.Names.GetNamedRange("GetSelectedLayerWorksheet");

                Common.Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
                Common.Globals_Accessor.TargetMachine = new TargetMachine("localhost");

                Group marsGroup = new Group("Mars", GroupType.ReferenceFrame, null);
                Group sunGroup = new Group("Sun", GroupType.ReferenceFrame, null);
                Group earthGroup = new Group("Earth", GroupType.ReferenceFrame, sunGroup);
                earthGroup.LayerIDs.Add("8e5cbbc4-9eb2-47e5-b3d7-7ca9babda477");
                sunGroup.Children.Add(earthGroup);

                LayerMap marsLayer = new LayerMap(emptyRange);
                marsLayer.MapType = LayerMapType.LocalInWWT;
                marsLayer.LayerDetails.Group = marsGroup;
                marsLayer.LayerDetails.ID = "MarsValidID1";

                LayerMap sunLayer = new LayerMap(emptyRange);
                sunLayer.MapType = LayerMapType.LocalInWWT;
                sunLayer.IsNotInSync = true;
                sunLayer.LayerDetails.Group = sunGroup;
                sunLayer.LayerDetails.ID = "sunLayerID1";

                WorkbookMap workbookMap = new WorkbookMap(null);
                workbookMap.AllLayerMaps.Add(marsLayer);
                workbookMap.AllLayerMaps.Add(sunLayer);

                Dictionary<string, string> actual = WorkbookMapExtensions.GetNamedRangesForInSyncLayers(workbookMap);
                Assert.IsTrue(actual.Count == 1);
                Assert.AreEqual("=GetSelectedLayerWorksheet!$A$1:$D$7", actual["GetSelectedLayerWorksheet"]);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion

        #region RemoveAffectedLayersTest

        /// <summary>
        /// A test for RemoveAffectedLayers
        /// </summary>
        [TestMethod()]
        public void RemoveAffectedLayersTest()
        {
            Group marsGroup = new Group("Mars", GroupType.ReferenceFrame, null);
            Group sunGroup = new Group("Sun", GroupType.ReferenceFrame, null);
            Group earthGroup = new Group("Earth", GroupType.ReferenceFrame, sunGroup);
            earthGroup.LayerIDs.Add("8e5cbbc4-9eb2-47e5-b3d7-7ca9babda477");
            sunGroup.Children.Add(earthGroup);

            LayerMap earthLayer = new LayerMap(new Layer());
            earthLayer.MapType = LayerMapType.WWT;
            earthLayer.LayerDetails.Group = earthGroup;
            earthLayer.LayerDetails.ID = "ValidID1";

            LayerMap marsLayer = new LayerMap(new Layer());
            marsLayer.MapType = LayerMapType.LocalInWWT;
            marsLayer.LayerDetails.Group = marsGroup;
            marsLayer.LayerDetails.ID = "MarsValidID1";

            LayerMap sunLayer = new LayerMap(new Layer());
            sunLayer.MapType = LayerMapType.LocalInWWT;
            sunLayer.LayerDetails.Group = sunGroup;
            sunLayer.LayerDetails.ID = "sunLayerID1";

            WorkbookMap workbookMap = new WorkbookMap(null);
            workbookMap.AllLayerMaps.Add(earthLayer);
            workbookMap.AllLayerMaps.Add(marsLayer);
            workbookMap.AllLayerMaps.Add(sunLayer);

            List<LayerMap> affectedLayerList = new List<LayerMap>();
            affectedLayerList.Add(marsLayer);
            affectedLayerList.Add(sunLayer);

            WorkbookMapExtensions.RemoveAffectedLayers(null, null);
            WorkbookMapExtensions.RemoveAffectedLayers(workbookMap, affectedLayerList);
            Assert.IsTrue(workbookMap.AllLayerMaps.Count == 1);
        }

        #endregion

        #region SerializeTest

        /// <summary>
        /// A test for Serialize
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void SerializeTest()
        {
            Application application = new Application();

            try
            {
                ThisAddIn_Accessor.ExcelApplication = application;
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);
                Name layerOne = book.Names.GetNamedRange("UpdateLayerRangeInitial");
                LayerMap layerMap = new LayerMap(layerOne);
                WorkbookMap workbookMap = new WorkbookMap(book);
                workbookMap.AllLayerMaps.Add(layerMap);
                string expected = "<?xml version=\"1.0\" encoding=\"utf-16\"?><WorkbookMap xmlns:d1p1=\"http://schemas.datacontract.org/2004/07/Microsoft.Research.Wwt.Excel.Addin\" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"Microsoft.Research.Wwt.Excel.Addin.WorkbooMap\"><d1p1:SerializableLayerMaps><d1p1:LayerMap><d1p1:LayerDetails xmlns:d4p1=\"http://schemas.datacontract.org/2004/07/Microsoft.Research.Wwt.Excel.Common\"><d4p1:AltColumn>-1</d4p1:AltColumn><d4p1:AltType>Depth</d4p1:AltType><d4p1:AltUnit>Meters</d4p1:AltUnit><d4p1:Color>ARGBColor:255:255:0:0</d4p1:Color><d4p1:ColorMapColumn>-1</d4p1:ColorMapColumn><d4p1:CoordinatesType>Spherical</d4p1:CoordinatesType><d4p1:DecColumn>-1</d4p1:DecColumn><d4p1:EndDateColumn>-1</d4p1:EndDateColumn><d4p1:EndTime>9999-12-31T23:59:59.9999999</d4p1:EndTime><d4p1:FadeSpan>PT0S</d4p1:FadeSpan><d4p1:FadeType>None</d4p1:FadeType><d4p1:GeometryColumn>-1</d4p1:GeometryColumn><d4p1:Group i:nil=\"true\" /><d4p1:HasTimeSeries>false</d4p1:HasTimeSeries><d4p1:ID i:nil=\"true\" /><d4p1:LatColumn>-1</d4p1:LatColumn><d4p1:LngColumn>-1</d4p1:LngColumn><d4p1:MarkerIndex>0</d4p1:MarkerIndex><d4p1:MarkerScale>World</d4p1:MarkerScale><d4p1:Name>UpdateLayerRangeInitial</d4p1:Name><d4p1:NameColumn>0</d4p1:NameColumn><d4p1:Opacity>1</d4p1:Opacity><d4p1:PlotType>Gaussian</d4p1:PlotType><d4p1:PointScaleType>StellarMagnitude</d4p1:PointScaleType><d4p1:RAColumn>-1</d4p1:RAColumn><d4p1:RAUnit>Hours</d4p1:RAUnit><d4p1:ReverseXAxis>false</d4p1:ReverseXAxis><d4p1:ReverseYAxis>false</d4p1:ReverseYAxis><d4p1:ReverseZAxis>false</d4p1:ReverseZAxis><d4p1:ScaleFactor>8</d4p1:ScaleFactor><d4p1:ShowFarSide>true</d4p1:ShowFarSide><d4p1:SizeColumn>-1</d4p1:SizeColumn><d4p1:StartDateColumn>-1</d4p1:StartDateColumn><d4p1:StartTime>0001-01-01T00:00:00</d4p1:StartTime><d4p1:TimeDecay>16</d4p1:TimeDecay><d4p1:Version>0</d4p1:Version><d4p1:XAxis>-1</d4p1:XAxis><d4p1:YAxis>-1</d4p1:YAxis><d4p1:ZAxis>-1</d4p1:ZAxis></d1p1:LayerDetails><d1p1:MapType>Local</d1p1:MapType><d1p1:MappedColumnType xmlns:d4p1=\"http://schemas.datacontract.org/2004/07/Microsoft.Research.Wwt.Excel.Common\"><d4p1:ColumnType>None</d4p1:ColumnType><d4p1:ColumnType>None</d4p1:ColumnType><d4p1:ColumnType>None</d4p1:ColumnType><d4p1:ColumnType>None</d4p1:ColumnType><d4p1:ColumnType>None</d4p1:ColumnType><d4p1:ColumnType>None</d4p1:ColumnType></d1p1:MappedColumnType><d1p1:RangeAddress>=UpdateLayer!$A$1:$F$9</d1p1:RangeAddress><d1p1:RangeDisplayName>UpdateLayerRangeInitial</d1p1:RangeDisplayName></d1p1:LayerMap></d1p1:SerializableLayerMaps><d1p1:SerializableSelectedLayerMap i:nil=\"true\" /></WorkbookMap>";
                string actual = WorkbookMapExtensions.Serialize(workbookMap);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion
    }
}
