//-----------------------------------------------------------------------
// <copyright file="LayerMapExtensionsTest.cs" company="Microsoft Corporation">
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
    /// This is a test class for LayerMapExtensionsTest and is intended to contain all LayerMapExtensionsTest Unit Tests
    /// </summary>
    [TestClass()]
    public class LayerMapExtensionsTest
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

        #region UpdateLayerMapPropertiesTest

        /// <summary>
        /// A test for UpdateLayerMapProperties
        /// </summary>
        [TestMethod()]
        public void UpdateLayerMapPropertiesTest()
        {
            Layer layer = new Layer();
            Layer dummyLayer = new Layer();
            LayerMap expected = new LayerMap(dummyLayer);

            // UpdateLayerMapProperties will set the layer in layer details.
            expected.UpdateLayerMapProperties(layer);

            // LayerDetails should not be dummy layer. It should be set with layer object.
            Assert.AreEqual(expected.LayerDetails, layer);
        }

        /// <summary>
        /// A test for UpdateLayerMapProperties
        /// </summary>
        [TestMethod()]
        public void UpdateLayerMapPropertiesNullTest()
        {
            LayerMap actual;
            actual = LayerMapExtensions.UpdateLayerMapProperties(null, null);
            Assert.IsNull(actual);
        }

        #endregion

        #region CanUpdateWWTTest

        /// <summary>
        /// A test for CanUpdateWWT
        /// </summary>
        [TestMethod()]
        public void CanUpdateWWTTestWWTOnly()
        {
            LayerMap selectedlayer = new LayerMap(new Layer());
            bool expected = true;
            bool actual;
            actual = LayerMapExtensions.CanUpdateWWT(selectedlayer);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for CanUpdateWWT
        /// </summary>
        [TestMethod()]
        public void CanUpdateWWTTestWWTLocalAndSync()
        {
            LayerMap selectedlayer = new LayerMap(new Layer());
            selectedlayer.MapType = LayerMapType.LocalInWWT;
            selectedlayer.IsNotInSync = false;
            bool expected = true;
            bool actual;
            actual = LayerMapExtensions.CanUpdateWWT(selectedlayer);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for CanUpdateWWT
        /// </summary>
        [TestMethod()]
        public void CanUpdateWWTTestWWTLocalAndNotSync()
        {
            LayerMap selectedlayer = new LayerMap(new Layer());
            selectedlayer.MapType = LayerMapType.LocalInWWT;
            selectedlayer.IsNotInSync = true;
            bool expected = false;
            bool actual;
            actual = LayerMapExtensions.CanUpdateWWT(selectedlayer);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for CanUpdateWWT
        /// </summary>
        [TestMethod()]
        public void CanUpdateWWTTestWWTLocalOnly()
        {
            LayerMap selectedlayer = new LayerMap(new Layer());
            selectedlayer.MapType = LayerMapType.Local;
            bool expected = false;
            bool actual;
            actual = LayerMapExtensions.CanUpdateWWT(selectedlayer);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for CanUpdateWWT
        /// </summary>
        [TestMethod()]
        public void CanUpdateWWTTestWWTNone()
        {
            LayerMap selectedlayer = new LayerMap(new Layer());
            selectedlayer.MapType = LayerMapType.None;
            bool expected = false;
            bool actual;
            actual = LayerMapExtensions.CanUpdateWWT(selectedlayer);
            Assert.AreEqual(expected, actual);
        }

        #endregion

        #region IsLayerCreatedTest

        /// <summary>
        /// A test for IsLayerCreated
        /// </summary>
        [TestMethod()]
        public void IsLayerCreatedTestNoLayerID()
        {
            LayerMap selectedLayerMap = new LayerMap(new Layer());
            bool expected = true;
            bool actual;
            actual = LayerMapExtensions.IsLayerCreated(selectedLayerMap);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsLayerCreated
        /// </summary>
        [TestMethod()]
        public void IsLayerCreatedTestNull()
        {
            bool expected = false;
            bool actual;
            actual = LayerMapExtensions.IsLayerCreated(null);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsLayerCreated
        /// </summary>
        [TestMethod()]
        public void IsLayerCreatedTestValidLayerID()
        {
            LayerMap selectedLayerMap = new LayerMap(new Layer());
            selectedLayerMap.LayerDetails.ID = System.Guid.NewGuid().ToString();
            bool expected = false;
            bool actual;
            actual = LayerMapExtensions.IsLayerCreated(selectedLayerMap);
            Assert.AreEqual(expected, actual);
        }

        #endregion

        #region UpdateHeaderPropertiesTest

        /// <summary>
        /// A test for UpdateHeaderProperties
        /// </summary>
        [TestCategory("Interactive"), TestMethod()]
        public void UpdateHeaderPropertiesTest()
        {
            Application application = new Application();
            try
            {
                Workbook book = application.OpenWorkbook("WorkbookTestData.xlsx", false);

                // Get the named range stored in the test data excel file.
                // This range refers to address "$A$1:$D$7".
                Name oldName = book.Names.GetNamedRange("GetSelectedLayerWorksheet");
                Name newName = book.Names.GetNamedRange("UpdateHeaderPropertiesTestRADEC");

                LayerMap selectedlayer = new LayerMap(oldName);

                Range selectedRange = newName.RefersToRange;
                LayerMapExtensions.UpdateHeaderProperties(selectedlayer, selectedRange);
                Assert.AreEqual(selectedlayer.MappedColumnType[0], ColumnType.RA);
                Assert.AreEqual(selectedlayer.MappedColumnType[1], ColumnType.Dec);
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for UpdateHeaderProperties
        /// </summary>
        [TestMethod()]
        public void UpdateHeaderPropertiesTestNull()
        {
            Application application = new Application();
            try
            {
                LayerMap selectedlayer = new LayerMap(new Layer());
                LayerMapExtensions.UpdateHeaderProperties(selectedlayer, null);
                LayerMapExtensions.UpdateHeaderProperties(null, null);
                Assert.IsNull(selectedlayer.RangeAddress);
            }
            finally
            {
                application.Close();
            }
        }

        #endregion

        #region UpdateMappedColumnsTest

        /// <summary>
        /// A test for UpdateMappedColumns
        /// </summary>
        [TestMethod()]
        public void UpdateMappedColumnsTestNull()
        {
            LayerMap selectedLayerMap = new LayerMap(new Layer());

            // Check what happens if we send null.
            LayerMapExtensions.UpdateMappedColumns(null);

            // Check what happens if we send null group.
            LayerMapExtensions.UpdateMappedColumns(selectedLayerMap);

            Assert.IsNull(selectedLayerMap.LayerDetails.Group);
        }

        /// <summary>
        /// A test for UpdateMappedColumns
        /// </summary>
        [TestMethod()]
        public void UpdateMappedColumnsEarthTest()
        {
            Group earthGroup = new Group("Earth", GroupType.ReferenceFrame, new Group("Sun", GroupType.ReferenceFrame, null));

            LayerMap selectedLayerMap = new LayerMap(new Layer());
            selectedLayerMap.LayerDetails.Group = earthGroup;

            Collection<ColumnType> mappedColTypes = new Collection<ColumnType>();
            mappedColTypes.Add(ColumnType.RA);
            mappedColTypes.Add(ColumnType.Dec);
            mappedColTypes.Add(ColumnType.Lat);
            mappedColTypes.Add(ColumnType.Long);

            selectedLayerMap.MappedColumnType = mappedColTypes;

            LayerMapExtensions.UpdateMappedColumns(selectedLayerMap);

            Assert.AreEqual(selectedLayerMap.MappedColumnType[0], ColumnType.None);
            Assert.AreEqual(selectedLayerMap.MappedColumnType[1], ColumnType.None);
            Assert.AreEqual(selectedLayerMap.MappedColumnType[2], ColumnType.Lat);
            Assert.AreEqual(selectedLayerMap.MappedColumnType[3], ColumnType.Long);
        }

        /// <summary>
        /// A test for UpdateMappedColumns
        /// </summary>
        [TestMethod()]
        public void UpdateMappedColumnsSkyTest()
        {
            Group skyGroup = new Group("Sky", GroupType.ReferenceFrame, null);

            LayerMap selectedLayerMap = new LayerMap(new Layer());
            selectedLayerMap.LayerDetails.Group = skyGroup;

            Collection<ColumnType> mappedColTypes = new Collection<ColumnType>();
            mappedColTypes.Add(ColumnType.RA);
            mappedColTypes.Add(ColumnType.Dec);
            mappedColTypes.Add(ColumnType.Lat);
            mappedColTypes.Add(ColumnType.Long);

            selectedLayerMap.MappedColumnType = mappedColTypes;

            LayerMapExtensions.UpdateMappedColumns(selectedLayerMap);

            Assert.AreEqual(selectedLayerMap.MappedColumnType[0], ColumnType.RA);
            Assert.AreEqual(selectedLayerMap.MappedColumnType[1], ColumnType.Dec);
            Assert.AreEqual(selectedLayerMap.MappedColumnType[2], ColumnType.None);
            Assert.AreEqual(selectedLayerMap.MappedColumnType[3], ColumnType.None);
        }

        #endregion

        #region GroupCollection

        /// <summary>
        /// A test for BuildGroupCollection
        /// </summary>
        [TestMethod()]
        public void BuildGroupCollectionTest()
        {
            Application application = new Application();

            try
            {
                Group parentGroup = new Group("Sun", GroupType.ReferenceFrame, null);
                Group childGroup = new Group("Earth", GroupType.ReferenceFrame, parentGroup);
                Layer layer = new Layer();
                layer.Name = "Layer1";
                layer.Group = childGroup;
                Layer layerMap = new Layer();
                layerMap.Name = "Layer2";
                layerMap.Group = new Group("Sun", GroupType.ReferenceFrame, null);
                LayerMap localLayerMap = new LayerMap(layer);
                localLayerMap.MapType = LayerMapType.Local;
                LayerMap wwtLayerMap = new LayerMap(layerMap);
                wwtLayerMap.MapType = LayerMapType.WWT;

                List<GroupChildren> wwtGroups = new List<GroupChildren>();
                List<GroupChildren> localGroups = new List<GroupChildren>();
                List<GroupChildren> existingGroups = new List<GroupChildren>();
                List<GroupChildren> groupwithChildren = new List<GroupChildren>();
                GroupChildren children = new GroupChildren();
                children.Group = parentGroup;
                GroupChildren children1 = new GroupChildren();
                children1.Group = childGroup;
                GroupChildren childNode = new GroupChildren();
                childNode.Group = parentGroup;
                childNode.Children.Add(children1);
                existingGroups.Add(children);
                groupwithChildren.Add(childNode);
                LayerMapExtensions.BuildGroupCollection(wwtLayerMap, wwtGroups);
                LayerMapExtensions.BuildGroupCollection(localLayerMap, localGroups);
                LayerMapExtensions.BuildGroupCollection(localLayerMap, existingGroups);
                LayerMapExtensions.BuildGroupCollection(localLayerMap, groupwithChildren);

                Assert.AreEqual(1, wwtGroups.Count);
                foreach (GroupChildren child in wwtGroups)
                {
                    Assert.AreEqual(1, child.AllChildren.Count);
                    Assert.AreEqual("Sun", child.Name); 
                }

                Assert.AreEqual(1, localGroups.Count);
                foreach (GroupChildren child in localGroups)
                {
                    Assert.AreEqual(1, child.Children.Count);
                    Assert.AreEqual("Sun", child.Name);
                }

                Assert.AreEqual(1, existingGroups.Count);
                foreach (GroupChildren child in existingGroups)
                {
                    Assert.AreEqual(1, child.Children.Count);
                    foreach (GroupChildren childrenVal in child.Children)
                    {
                        Assert.AreEqual("Earth", childrenVal.Name);
                    }
                }

                Assert.AreEqual(1, groupwithChildren.Count);
                foreach (GroupChildren child in groupwithChildren)
                {
                    Assert.AreEqual(1, child.Children.Count);
                    foreach (GroupChildren childrenVal in child.Children)
                    {
                        Assert.AreEqual(1, childrenVal.AllChildren.Count);
                    }
                }
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for AddChildNodesToGroup
        /// </summary>
        [TestMethod()]
        public void AddChildNodesToGroupTest()
        {
            Application application = new Application();

            try
            {
                // Where ever we are creating object for WorkflowController_Accessor, we need to set the ThisAddIn_Accessor.ExcelApplication to 
                // wither null or actual application object.
                ThisAddIn_Accessor.ExcelApplication = application; 

                Group childGroup = new Group("Earth", GroupType.ReferenceFrame, null);
                Layer layer = new Layer();
                layer.Name = "Layer1";
                layer.Group = childGroup;
                LayerMap_Accessor localLayerMap = new LayerMap_Accessor(layer);
                localLayerMap.MapType = LayerMapType.Local;

                GroupChildren groupChild = LayerMapExtensions_Accessor.AddChildNodesToGroup(localLayerMap);
                Assert.AreEqual("Earth", groupChild.Name);
                Assert.AreEqual(1, groupChild.AllChildren.Count);
                Assert.AreEqual(1, groupChild.Layers.Count);
                foreach (Layer layerVal in groupChild.Layers)
                {
                    Assert.AreEqual("Layer1", layerVal.Name);
                }
            }
            finally
            {
                application.Close();
            }
        }

        /// <summary>
        /// A test for AddLayerNode
        /// </summary>
        [TestMethod()]
        public void AddLayerNodeTest()
        {
            Application application = new Application();

            try
            {
                // Where ever we are creating object for WorkflowController_Accessor, we need to set the ThisAddIn_Accessor.ExcelApplication to 
                // wither null or actual application object.
                ThisAddIn_Accessor.ExcelApplication = application; 
                List<GroupChildren> groups = new List<GroupChildren>();
                List<GroupChildren> nestedGroups = new List<GroupChildren>();
                
                Group parentGroup = new Group("Earth", GroupType.ReferenceFrame, null);
                Layer layer = new Layer();
                layer.Name = "Layer1";
                layer.Group = parentGroup;
                
                LayerMap_Accessor localLayerMap = new LayerMap_Accessor(layer);
                localLayerMap.MapType = LayerMapType.Local;
                localLayerMap.RangeDisplayName = "Sheet_1";
                
                Group childGroup = new Group("Earth", GroupType.ReferenceFrame, parentGroup);
                Layer nestedLayer = new Layer();
                nestedLayer.Name = "Layer1";
                nestedLayer.Group = childGroup;
                
                LayerMap_Accessor nestedLayerMap = new LayerMap_Accessor(nestedLayer);
                nestedLayerMap.MapType = LayerMapType.Local;
                nestedLayerMap.RangeDisplayName = "Sheet_1";

                LayerMapExtensions_Accessor.AddLayerNode(groups, localLayerMap);
                LayerMapExtensions_Accessor.AddLayerNode(nestedGroups, nestedLayerMap);

                Assert.AreEqual(1, groups.Count);
                foreach (GroupChildren group in groups)
                {
                    Assert.AreEqual("Earth", group.Name);
                    Assert.AreEqual(1, group.AllChildren.Count);
                    foreach (LayerMap layerVal in group.AllChildren)
                    {
                        Assert.AreEqual("Layer1", layerVal.LayerDetails.Name);
                    }
                }

                Assert.AreEqual(1, nestedGroups.Count);
                foreach (GroupChildren group in nestedGroups)
                {
                    Assert.AreEqual("Earth", group.Name);
                    Assert.AreEqual(1, group.Children.Count);
                }
            }
            finally
            {
                application.Close();
            }
        }
        #endregion
    }
}
