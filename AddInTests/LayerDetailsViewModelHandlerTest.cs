//-----------------------------------------------------------------------
// <copyright file="LayerDetailsViewModelHandlerTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.ObjectModel;
using System.Linq;
using Microsoft.Research.Wwt.Excel.Addin;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{
    /// <summary>
    /// This is a test class for LayerDetailsViewModelHandlerTest and is intended to contain all LayeLayerDetailsViewModelHandlerTestrDetailsViewModelTest Unit Tests
    /// </summary>
    [TestClass()]
    public class LayerDetailsViewModelHandlerTest
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

        #region Layer Selection Handler
        /// <summary>
        /// A test for LayerSelectionHandler Constructor
        /// </summary>
        [TestMethod()]
        public void LayerSelectionHandlerConstructorTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            LayerDetailsViewModel_Accessor.LayerSelectionHandler target = new LayerDetailsViewModel_Accessor.LayerSelectionHandler(layerDetailsViewModel);
            Assert.AreEqual(target.parent, layerDetailsViewModel);
        }

        /// <summary>
        /// A test for LayerSelectionHandler Execute
        /// </summary>
        [TestMethod()]
        public void LayerSelectionHandlerExecuteTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            layerDetailsViewModel.LayerSelectionChangedEvent += new EventHandler(LayerModelLayerSelectionChangedEvent);

            LayerMapDropDownViewModel layerMapDropDown = new LayerMapDropDownViewModel();
            layerMapDropDown.ID = "1";
            layerMapDropDown.Name = "Select One";

            Layer layer = new Layer();
            layer.Name = "Layer1";

            LayerMap layerMap = new LayerMap(layer);
            layerMap.MapType = LayerMapType.Local;
            layerMap.HeaderRowData = new Collection<string>();
            layerMap.MappedColumnType = new Collection<ColumnType>();
            layerDetailsViewModel.Currentlayer = layerMap;

            LayerDetailsViewModel_Accessor.LayerSelectionHandler target = new LayerDetailsViewModel_Accessor.LayerSelectionHandler(layerDetailsViewModel);
            target.Execute(layerMapDropDown);
            Assert.IsNull(layerDetailsViewModel.Currentlayer);
        }

        #endregion

        #region Control State ChangeHandler

        /// <summary>
        /// A test for ControlStateChangeHandler Constructor
        /// </summary>
        [TestMethod()]
        public void ControlStateChangeHandlerConstructorTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            LayerDetailsViewModel_Accessor.ControlStateChangeHandler target = new LayerDetailsViewModel_Accessor.ControlStateChangeHandler(layerDetailsViewModel);
            Assert.AreEqual(target.parent, layerDetailsViewModel);
        }

        /// <summary>
        /// A test for Execute ControlStateChangeHandler
        /// </summary>
        [TestMethod()]
        public void ControlStateChangeHandlerExecuteTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            layerDetailsViewModel.CustomTaskPaneStateChangedEvent += new EventHandler(LayerModelCustomTaskPaneStateChangedEvent);
            LayerDetailsViewModel_Accessor.ControlStateChangeHandler target = new LayerDetailsViewModel_Accessor.ControlStateChangeHandler(layerDetailsViewModel);
            layerDetailsViewModel.LayerDataDisplayName = "Layer1";
            target.Execute(layerDetailsViewModel);
            Assert.AreEqual(layerDetailsViewModel.LayerDataDisplayName, "Layer2");
        }

        #endregion

        #region View in WWT Handler

        /// <summary>
        /// A test for ViewInWWTHandler Constructor
        /// </summary>
        [TestMethod()]
        public void LayerModelViewInWWTHandlerConstructorTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            LayerDetailsViewModel_Accessor.ViewInWWTHandler target = new LayerDetailsViewModel_Accessor.ViewInWWTHandler(layerDetailsViewModel);
            Assert.AreEqual(target.parent, layerDetailsViewModel);
        }

        /// <summary>
        /// A test for Execute ViewInWWTHandler
        /// </summary>
        [TestMethod()]
        public void LayerModelViewInWWTHandlerExecuteTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            layerDetailsViewModel.ViewnInWWTClickedEvent += new EventHandler(LayerModelViewnInWWTClickedEvent);
            LayerDetailsViewModel_Accessor.ViewInWWTHandler target = new LayerDetailsViewModel_Accessor.ViewInWWTHandler(layerDetailsViewModel);
            layerDetailsViewModel.LayerDataDisplayName = "Layer1";
            target.Execute(layerDetailsViewModel);
            Assert.AreEqual(layerDetailsViewModel.LayerDataDisplayName, "Layer2");
        }

        #endregion

        #region Show Range Handler
        /// <summary>
        /// A test for ShowRangeHandler Constructor
        /// </summary>
        [TestMethod()]
        public void LayerModelShowRangeHandlerConstructorTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            LayerDetailsViewModel_Accessor.ShowRangeHandler target = new LayerDetailsViewModel_Accessor.ShowRangeHandler(layerDetailsViewModel);
            Assert.AreEqual(target.parent, layerDetailsViewModel);
        }

        /// <summary>
        /// A test for Execute ShowRangeHandler
        /// </summary>
        [TestMethod()]
        public void LayerModelShowRangeHandlerExecuteTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            layerDetailsViewModel.ShowRangeClickedEvent += new EventHandler(LayerModelShowRangeClickedEvent);
            LayerDetailsViewModel_Accessor.ShowRangeHandler target = new LayerDetailsViewModel_Accessor.ShowRangeHandler(layerDetailsViewModel);
            layerDetailsViewModel.LayerDataDisplayName = "Layer1";
            target.Execute(layerDetailsViewModel);
            Assert.AreEqual(layerDetailsViewModel.LayerDataDisplayName, "Layer2");
        }

        #endregion

        #region Delete mapping

        /// <summary>
        /// A test for DeleteMappingHandler Constructor
        /// </summary>
        [TestMethod()]
        public void LayerModelDeleteMappingHandlerConstructorTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            LayerDetailsViewModel_Accessor.DeleteMappingHandler target = new LayerDetailsViewModel_Accessor.DeleteMappingHandler(layerDetailsViewModel);
            Assert.AreEqual(target.parent, layerDetailsViewModel);
        }

        /// <summary>
        /// A test for Execute
        /// </summary>
        [TestMethod()]
        public void LayerModelDeleteMappingExecuteTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            layerDetailsViewModel.DeleteMappingClickedEvent += new EventHandler(LayerModelDeleteMappingClickedEvent);

            LayerDetailsViewModel_Accessor.DeleteMappingHandler target = new LayerDetailsViewModel_Accessor.DeleteMappingHandler(layerDetailsViewModel);
            layerDetailsViewModel.LayerDataDisplayName = "Layer1";
            target.Execute(layerDetailsViewModel);
            Assert.AreEqual(layerDetailsViewModel.LayerDataDisplayName, "Layer2");
        }

        #endregion

        #region Update Layer Handler

        /// <summary>
        /// A test for UpdateLayerHandler Constructor
        /// </summary>
        [TestMethod()]
        public void LayerModelUpdateLayerHandlerConstructorTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            LayerDetailsViewModel_Accessor.UpdateLayerHandler target = new LayerDetailsViewModel_Accessor.UpdateLayerHandler(layerDetailsViewModel);
            Assert.AreEqual(target.parent, layerDetailsViewModel);
        }

        /// <summary>
        /// A test for Execute Update layer
        /// </summary>
        [TestMethod()]
        public void LayerModelUpdateLayerExecuteTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            layerDetailsViewModel.UpdateLayerClickedEvent += new EventHandler(LayerModelUpdateLayerClickedEvent);

            LayerDetailsViewModel_Accessor.UpdateLayerHandler target = new LayerDetailsViewModel_Accessor.UpdateLayerHandler(layerDetailsViewModel);
            layerDetailsViewModel.LayerDataDisplayName = "Layer1";
            target.Execute(layerDetailsViewModel);
            Assert.AreEqual(layerDetailsViewModel.LayerDataDisplayName, "Layer2");
        }
        #endregion

        #region Layer Map Name Change Handler
        /// <summary>
        /// A test for LayerMapNameChangeHandler Constructor
        /// </summary>
        [TestMethod()]
        public void LayerModelLayerMapNameChangeHandlerConstructorTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            LayerDetailsViewModel_Accessor.LayerMapNameChangeHandler target = new LayerDetailsViewModel_Accessor.LayerMapNameChangeHandler(layerDetailsViewModel);
            Assert.AreEqual(target.parent, layerDetailsViewModel);
        }

        /// <summary>
        /// A test for Execute layer display name
        /// </summary>
        [TestMethod()]
        public void LayerModelLayerDisplayExecuteTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            LayerDetailsViewModel_Accessor.LayerMapNameChangeHandler target = new LayerDetailsViewModel_Accessor.LayerMapNameChangeHandler(layerDetailsViewModel);
            string layerName = "Layer1";
            Layer layer = new Layer();
            layer.Name = layerName;

            LayerMap layerMap = new LayerMap(layer);
            layerMap.MapType = LayerMapType.Local;
            layerMap.HeaderRowData = new Collection<string>();
            layerMap.MappedColumnType = new Collection<ColumnType>();
            layerDetailsViewModel.Currentlayer = layerMap;

            target.Execute(layerName);
            Assert.AreEqual(layerDetailsViewModel.SelectedLayerName, layerName);
        }
        #endregion

        #region Get Layer Data Handler

        /// <summary>
        /// A test for GetLayerDataHandler Constructor
        /// </summary>
        [TestMethod()]
        public void LayerModelGetLayerDataHandlerConstructorTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            LayerDetailsViewModel_Accessor.GetLayerDataHandler target = new LayerDetailsViewModel_Accessor.GetLayerDataHandler(layerDetailsViewModel);
            Assert.AreEqual(target.parent, layerDetailsViewModel);
        }

        /// <summary>
        /// A test for Execute get layer data
        /// </summary>
        [TestMethod()]
        public void LayerModelGetLayerDataExecuteTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            layerDetailsViewModel.GetLayerDataClickedEvent += new EventHandler(LayerModelGetLayerDataClickedEvent);

            LayerDetailsViewModel_Accessor.GetLayerDataHandler target = new LayerDetailsViewModel_Accessor.GetLayerDataHandler(layerDetailsViewModel);
            layerDetailsViewModel.LayerDataDisplayName = "Layer1";
            target.Execute(layerDetailsViewModel);
            Assert.AreEqual(layerDetailsViewModel.LayerDataDisplayName, "Layer2");
        }

        #endregion

        #region Fade Time Change Handler

        /// <summary>
        /// A test for FadeTimeChangeHandler Constructor
        /// </summary>
        [TestMethod()]
        public void LayerModelFadeTimeChangeHandlerConstructorTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            LayerDetailsViewModel_Accessor.FadeTimeChangeHandler target = new LayerDetailsViewModel_Accessor.FadeTimeChangeHandler(layerDetailsViewModel);
            Assert.AreEqual(target.parent, layerDetailsViewModel);
        }

        /// <summary>
        /// A test for fade time Execute
        /// </summary>
        [TestMethod()]
        public void LayerModelFadeTimeExecuteTest()
        {
            LayerDetailsViewModel layerDetailsViewModel = new LayerDetailsViewModel();
            Layer layer = new Layer();
            layer.Name = "Layer1";

            LayerMap layerMap = new LayerMap(layer);
            layerMap.MapType = LayerMapType.Local;
            layerMap.HeaderRowData = new Collection<string>();
            layerMap.MappedColumnType = new Collection<ColumnType>();
            layerDetailsViewModel.Currentlayer = layerMap;
            LayerDetailsViewModel_Accessor.FadeTimeChangeHandler target = new LayerDetailsViewModel_Accessor.FadeTimeChangeHandler(layerDetailsViewModel);
            string fadeTime = "10:10:10";
            target.Execute(fadeTime);
            Assert.AreEqual(fadeTime, layerDetailsViewModel.FadeTime.ToString());
        }

        #endregion

        #region Events Tests

        /// <summary>
        /// A test for OnMapColumnSelectionChanged for Mag
        /// </summary>
        [TestMethod()]
        public void OnMapColumnSelectionChangedEventMagTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();

            ColumnViewModel columnView = new ColumnViewModel();
            columnView.ExcelHeaderColumn = "LAT";
            columnView.WWTColumns = new ObservableCollection<Column>();
            ColumnExtensions.PopulateColumnList().ToList().ForEach(col => columnView.WWTColumns.Add(col));
            columnView.SelectedWWTColumn = columnView.WWTColumns.Where(column => column.ColType == ColumnType.Mag).FirstOrDefault();

            target.ColumnsView = new ObservableCollection<ColumnViewModel>();
            target.ColumnsView.Add(columnView);

            target.sizeColumnList = new ObservableCollection<System.Collections.Generic.KeyValuePair<int, string>>();
            target.sizeColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(0, "LAT"));
            target.sizeColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(1, "MAG"));
            target.sizeColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(2, "LONG"));
            target.sizeColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(3, "DEPTH"));

            target.hoverTextColumnList = new ObservableCollection<System.Collections.Generic.KeyValuePair<int, string>>();
            target.hoverTextColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(0, "LAT"));
            target.hoverTextColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(1, "MAG"));
            target.hoverTextColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(2, "LONG"));
            target.hoverTextColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(3, "DEPTH"));

            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.StartTime = DateTime.Now;

            LayerMap_Accessor localLayerMap = new LayerMap_Accessor(layer);
            localLayerMap.MapType = LayerMapType.Local;
            localLayerMap.RangeDisplayName = "Sheet_1";
            target.currentLayer = localLayerMap;

            EventArgs e = null;
            target.OnMapColumnSelectionChanged(columnView, e);

            Assert.AreEqual(target.selectedSizeColumn.Key, 1);
            Assert.AreEqual(true, target.isMarkerTabEnabled);
        }

        /// <summary>
        /// A test for OnMapColumnSelectionChanged for depth
        /// </summary>
        [TestMethod()]
        public void OnMapColumnSelectionChangedEventDepthTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();

            ColumnViewModel columnView = new ColumnViewModel();
            columnView.ExcelHeaderColumn = "Depth";
            columnView.WWTColumns = new ObservableCollection<Column>();
            ColumnExtensions.PopulateColumnList().ToList().ForEach(col => columnView.WWTColumns.Add(col));
            columnView.SelectedWWTColumn = columnView.WWTColumns.Where(column => column.ColType == ColumnType.Depth).FirstOrDefault();

            target.ColumnsView = new ObservableCollection<ColumnViewModel>();
            target.ColumnsView.Add(columnView);

            target.sizeColumnList = new ObservableCollection<System.Collections.Generic.KeyValuePair<int, string>>();
            target.sizeColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(0, "LAT"));
            target.sizeColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(1, "MAG"));
            target.sizeColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(2, "LONG"));
            target.sizeColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(3, "DEPTH"));

            target.hoverTextColumnList = new ObservableCollection<System.Collections.Generic.KeyValuePair<int, string>>();
            target.hoverTextColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(0, "LAT"));
            target.hoverTextColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(1, "MAG"));
            target.hoverTextColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(2, "LONG"));
            target.hoverTextColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(3, "DEPTH"));

            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.StartTime = DateTime.Now;

            LayerMap_Accessor localLayerMap = new LayerMap_Accessor(layer);
            localLayerMap.MapType = LayerMapType.Local;
            localLayerMap.RangeDisplayName = "Sheet_1";
            target.currentLayer = localLayerMap;

            EventArgs e = null;
            target.OnMapColumnSelectionChanged(columnView, e);

            Assert.AreEqual(target.selectedSizeColumn.Key, 0);
            Assert.AreEqual(true, target.isMarkerTabEnabled);
            Assert.AreEqual(true, target.isDistanceVisible);
            Assert.AreEqual(false, target.IsRAUnitVisible);
            Assert.AreEqual(AngleUnit.Hours, target.selectedRAUnit.Key);
        }

        /// <summary>
        /// A test for OnMapColumnSelectionChanged for RA
        /// </summary>
        [TestMethod()]
        public void OnMapColumnSelectionChangedEventRAGeoTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();

            ColumnViewModel columnView = new ColumnViewModel();
            columnView.ExcelHeaderColumn = "RA";
            columnView.WWTColumns = new ObservableCollection<Column>();
            ColumnExtensions.PopulateColumnList().ToList().ForEach(col => columnView.WWTColumns.Add(col));
            columnView.SelectedWWTColumn = columnView.WWTColumns.Where(column => column.ColType == ColumnType.RA).FirstOrDefault();

            target.ColumnsView = new ObservableCollection<ColumnViewModel>();
            target.ColumnsView.Add(columnView);

            target.sizeColumnList = new ObservableCollection<System.Collections.Generic.KeyValuePair<int, string>>();
            target.sizeColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(0, "LAT"));
            target.sizeColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(1, "MAG"));
            target.sizeColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(2, "LONG"));
            target.sizeColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(3, "DEPTH"));

            target.hoverTextColumnList = new ObservableCollection<System.Collections.Generic.KeyValuePair<int, string>>();
            target.hoverTextColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(0, "LAT"));
            target.hoverTextColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(1, "MAG"));
            target.hoverTextColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(2, "LONG"));
            target.hoverTextColumnList.Add(new System.Collections.Generic.KeyValuePair<int, string>(3, "DEPTH"));

            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.StartTime = DateTime.Now;

            LayerMap_Accessor localLayerMap = new LayerMap_Accessor(layer);
            localLayerMap.MapType = LayerMapType.Local;
            localLayerMap.RangeDisplayName = "Sheet_1";
            target.currentLayer = localLayerMap;

            EventArgs e = null;
            target.OnMapColumnSelectionChanged(columnView, e);

            Assert.AreEqual(true, target.isMarkerTabEnabled);
            Assert.AreEqual(false, target.isDistanceVisible);
            Assert.AreEqual(true, target.IsRAUnitVisible);
            Assert.AreEqual(AngleUnit.Hours, target.selectedRAUnit.Key);

            columnView.SelectedWWTColumn = columnView.WWTColumns.Where(column => column.ColType == ColumnType.Geo).FirstOrDefault();

            target.ColumnsView = new ObservableCollection<ColumnViewModel>();
            target.ColumnsView.Add(columnView);

            target.OnMapColumnSelectionChanged(columnView, e);
            Assert.AreEqual(true, target.isMarkerTabEnabled);
            Assert.AreEqual(false, target.isDistanceVisible);
            Assert.AreEqual(false, target.IsRAUnitVisible);
        }

        /// <summary>
        /// A test for OnGroupSelectionChanged
        /// </summary>
        [TestMethod()]
        public void OnGroupSelectionChangedTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Group group = new Group("Sun", GroupType.ReferenceFrame, null);
            Group skyGroup = new Group("Sky", GroupType.ReferenceFrame, null);
            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.StartTime = DateTime.Now;

            LayerMap_Accessor localLayerMap = new LayerMap_Accessor(layer);
            localLayerMap.MapType = LayerMapType.Local;
            localLayerMap.RangeDisplayName = "Sheet_1";
            localLayerMap.HeaderRowData = new Collection<string>();
            localLayerMap.HeaderRowData.Add("Lat");

            localLayerMap.MappedColumnType = new Collection<ColumnType>();
            localLayerMap.MappedColumnType.Add(ColumnType.RA);
            target.currentLayer = localLayerMap;

            ColumnViewModel columnView = new ColumnViewModel();
            columnView.ExcelHeaderColumn = "RA";
            columnView.WWTColumns = new ObservableCollection<Column>();
            ColumnExtensions.PopulateColumnList().ToList().ForEach(col => columnView.WWTColumns.Add(col));
            columnView.SelectedWWTColumn = columnView.WWTColumns.Where(column => column.ColType == ColumnType.RA).FirstOrDefault();

            target.ColumnsView = new ObservableCollection<ColumnViewModel>();
            target.ColumnsView.Add(columnView);

            target.selectedScaleType = new System.Collections.Generic.KeyValuePair<ScaleType, string>(ScaleType.StellarMagnitude, "StellarMagnitude");

            object sender = group;
            EventArgs e = null;
            target.OnGroupSelectionChanged(sender, e);

            Assert.AreEqual(false, target.isRAUnitVisible);
            Assert.AreEqual("Sun", target.selectedGroupText);
            Assert.AreEqual(ScaleType.Power, target.selectedScaleType.Key);

            sender = skyGroup;
            target.OnGroupSelectionChanged(sender, e);
            Assert.AreEqual(ScaleType.StellarMagnitude, target.selectedScaleType.Key);
        }

        /// <summary>
        /// A test for OnLayerSelectionChangedEvent
        /// </summary>
        [TestMethod()]
        public void OnLayerSelectionChangedEventTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.StartTime = DateTime.Now;

            LayerMap_Accessor localLayerMap = new LayerMap_Accessor(layer);
            localLayerMap.MapType = LayerMapType.Local;
            localLayerMap.RangeDisplayName = "Sheet_1";
            localLayerMap.HeaderRowData = new Collection<string>();
            localLayerMap.HeaderRowData.Add("Lat");

            localLayerMap.MappedColumnType = new Collection<ColumnType>();
            localLayerMap.MappedColumnType.Add(ColumnType.RA);
            target.currentLayer = localLayerMap;

            object sender = localLayerMap;
            EventArgs e = null;
            target.OnLayerSelectionChangedEvent(sender, e);

            Assert.AreEqual("Layer1", target.SelectedLayerText);
        }

        #endregion Events Tests

        #region Private Events

        /// <summary>
        /// On custom state changed on the custom task pane
        /// </summary>
        /// <param name="sender">Layer view model</param>
        /// <param name="e">Routed event</param>
        private void LayerModelCustomTaskPaneStateChangedEvent(object sender, EventArgs e)
        {
            LayerDetailsViewModel layerViewModel = sender as LayerDetailsViewModel;
            Assert.AreEqual(layerViewModel.LayerDataDisplayName, "Layer1");
            layerViewModel.LayerDataDisplayName = "Layer2";
        }

        /// <summary>
        /// Event is fired on the delete mapping 
        /// </summary>
        /// <param name="sender">Layer details view model</param>
        /// <param name="e">Routed event</param>
        private void LayerModelDeleteMappingClickedEvent(object sender, EventArgs e)
        {
            LayerDetailsViewModel layerViewModel = sender as LayerDetailsViewModel;
            Assert.AreEqual(layerViewModel.LayerDataDisplayName, "Layer1");
            layerViewModel.LayerDataDisplayName = "Layer2";
        }

        /// <summary>
        /// Event is fired on the show range is clicked
        /// </summary>
        /// <param name="sender">Layer view model</param>
        /// <param name="e">Routed event</param>
        private void LayerModelShowRangeClickedEvent(object sender, EventArgs e)
        {
            LayerDetailsViewModel layerViewModel = sender as LayerDetailsViewModel;
            Assert.AreEqual(layerViewModel.LayerDataDisplayName, "Layer1");
            layerViewModel.LayerDataDisplayName = "Layer2";
        }

        /// <summary>
        /// Event fired on view in WWT clicked
        /// </summary>
        /// <param name="sender">Layer view model</param>
        /// <param name="e">Routed event</param>
        private void LayerModelViewnInWWTClickedEvent(object sender, EventArgs e)
        {
            LayerDetailsViewModel layerViewModel = sender as LayerDetailsViewModel;
            Assert.AreEqual(layerViewModel.LayerDataDisplayName, "Layer1");
            layerViewModel.LayerDataDisplayName = "Layer2";
        }

        /// <summary>
        /// Layer selection changed event
        /// </summary>
        /// <param name="sender">Layer map dropdown</param>
        /// <param name="e">Routed event</param>
        private void LayerModelLayerSelectionChangedEvent(object sender, EventArgs e)
        {
            LayerMap currentlayer = sender as LayerMap;

            // It is expected to be null
            Assert.IsNull(currentlayer);
        }

        /// <summary>
        /// Layer update layer event
        /// </summary>
        /// <param name="sender">Update layer</param>
        /// <param name="e">Routed event</param>
        private void LayerModelUpdateLayerClickedEvent(object sender, EventArgs e)
        {
            LayerDetailsViewModel layerViewModel = sender as LayerDetailsViewModel;
            Assert.AreEqual(layerViewModel.LayerDataDisplayName, "Layer1");
            layerViewModel.LayerDataDisplayName = "Layer2";
        }

        /// <summary>
        /// Get layer clicked
        /// </summary>
        /// <param name="sender">get layer data</param>
        /// <param name="e">Routed event</param>
        private void LayerModelGetLayerDataClickedEvent(object sender, EventArgs e)
        {
            LayerDetailsViewModel layerViewModel = sender as LayerDetailsViewModel;
            Assert.AreEqual(layerViewModel.LayerDataDisplayName, "Layer1");
            layerViewModel.LayerDataDisplayName = "Layer2";
        }

        #endregion Private Events
    }
}
