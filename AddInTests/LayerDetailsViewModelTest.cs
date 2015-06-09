//-----------------------------------------------------------------------
// <copyright file="LayerDetailsViewModelTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media;
using Microsoft.Research.Wwt.Excel.Addin;
using Microsoft.Research.Wwt.Excel.Addin.Properties;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{
    /// <summary>
    /// This is a test class for LayerDetailsViewModelTest and is intended to contain all LayerDetailsViewModelTest Unit Tests
    /// </summary>
    [TestClass()]
    public class LayerDetailsViewModelTest
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
        /// A test for LayerDataDisplayName
        /// </summary>
        [TestMethod()]
        public void LayerDataDisplayNameTest()
        {
            string actual = "MyLayerDataDisplayName";
            LayerDetailsViewModel target = new LayerDetailsViewModel();
            target.LayerDataDisplayName = actual;

            string expected = target.LayerDataDisplayName;

            Assert.AreEqual(actual, expected);
        }

        /// <summary>
        /// A test for get BeginDate
        /// </summary>
        [TestMethod()]
        public void BeginDateGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.StartTime = DateTime.Now;

            LayerMap_Accessor localLayerMap = new LayerMap_Accessor(layer);
            localLayerMap.MapType = LayerMapType.Local;
            localLayerMap.RangeDisplayName = "Sheet_1";
            target.currentLayer = localLayerMap;
            DateTime actual = target.BeginDate;
            Assert.IsNotNull(actual);
            Assert.AreEqual(layer.StartTime, actual);
        }

        /// <summary>
        /// A test method for getting min value for begin date
        /// </summary>
        [TestMethod()]
        public void BeginDateGetMinValueTest()
        {
            LayerDetailsViewModel target = new LayerDetailsViewModel();
            DateTime actual = target.BeginDate;
            DateTime expected = DateTime.MinValue;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for set BeginDate
        /// </summary>
        [TestMethod()]
        public void BeginDateSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.StartTime = DateTime.Now;

            LayerMap_Accessor localLayerMap = new LayerMap_Accessor(layer);
            localLayerMap.MapType = LayerMapType.Local;
            localLayerMap.RangeDisplayName = "Sheet_1";
            target.currentLayer = localLayerMap;
            DateTime expected = DateTime.Now;
            target.BeginDate = expected;
            Assert.AreEqual(expected, target.BeginDate);
        }

        /// <summary>
        /// A test for get ColorBackground
        /// </summary>
        [TestMethod()]
        public void ColorBackgroundGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.Color = "ARGBColor:255:255:255:255";

            LayerMap_Accessor localLayerMap = new LayerMap_Accessor(layer);
            localLayerMap.MapType = LayerMapType.Local;
            localLayerMap.RangeDisplayName = "Sheet_1";
            target.currentLayer = localLayerMap;
            SolidColorBrush actual = (SolidColorBrush)target.ColorBackground;
            SolidColorBrush expected = new SolidColorBrush(Color.FromArgb(255, 255, 255, 255));
            Assert.AreEqual(expected.Color, actual.Color);
        }

        /// <summary>
        /// A test for get null ColorBackground
        /// </summary>
        [TestMethod()]
        public void ColorBackgroundNullGetTest()
        {
            LayerDetailsViewModel target = new LayerDetailsViewModel();
            SolidColorBrush expected = null;
            Assert.AreEqual(expected, target.ColorBackground);
        }

        /// <summary>
        /// A Test for set ColorBackground
        /// </summary>
        [TestMethod()]
        public void ColorBackgroundSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.Color = "ARGBColor:255:255:255:255";

            LayerMap_Accessor localLayerMap = new LayerMap_Accessor(layer);
            localLayerMap.MapType = LayerMapType.Local;
            localLayerMap.RangeDisplayName = "Sheet_1";
            target.currentLayer = localLayerMap;
            SolidColorBrush expected = new SolidColorBrush(Color.FromArgb(255, 255, 255, 255));
            target.ColorBackground = expected;
            Assert.AreEqual(expected.Color, ((SolidColorBrush)target.ColorBackground).Color);
        }

        /// <summary>
        ///  A Test for get Columns
        /// </summary>
        [TestMethod()]
        public void ColumnsViewGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.column = new ObservableCollection<ColumnViewModel>();
            target.column.Add(new ColumnViewModel() { ExcelHeaderColumn = "RA", SelectedWWTColumn = new Column(ColumnType.Alt, "RA", new Collection<string>() { "RA" }) });
            ObservableCollection<ColumnViewModel> actual = target.ColumnsView;
            Assert.IsNotNull(actual);
            Assert.AreEqual(target.column, actual);
        }

        /// <summary>
        /// A Test for set ColumnsView
        /// </summary>
        [TestMethod()]
        public void ColumnsViewSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            ObservableCollection<ColumnViewModel> expected = new ObservableCollection<ColumnViewModel>();
            expected.Add(new ColumnViewModel() { ExcelHeaderColumn = "RA", SelectedWWTColumn = new Column(ColumnType.Alt, "RA", new Collection<string>() { "RA" }) });
            target.ColumnsView = expected;
            if (target.ColumnsView == null || target.ColumnsView.Count == 0)
            {
                Assert.Fail("Columns view not set.");
            }
            Assert.AreEqual(expected.Count, target.ColumnsView.Count);
        }

        /// <summary>
        /// A Test for ConvertToSolidColorBrush
        /// </summary>
        [TestMethod()]
        public void ConvertToSolidColorBrushTest()
        {
            string colorArgb = "ARGBColor:255:255:255:255";
            SolidColorBrush actual = LayerDetailsViewModel.ConvertToSolidColorBrush(colorArgb);
            SolidColorBrush expected = new SolidColorBrush(Color.FromArgb(255, 255, 255, 255));
            Assert.AreEqual(expected.Color, actual.Color);
        }

        /// <summary>
        /// A Test for default ConvertToSolidColorBrush
        /// </summary>
        [TestMethod()]
        public void ConvertToDefaultSolidColorBrushTest()
        {
            string colorArgb = string.Empty;
            SolidColorBrush actual = LayerDetailsViewModel.ConvertToSolidColorBrush(colorArgb);
            SolidColorBrush expected = new SolidColorBrush(System.Windows.Media.Color.FromArgb(System.Drawing.Color.Red.A, System.Drawing.Color.Red.R, System.Drawing.Color.Red.G, System.Drawing.Color.Red.B));
            Assert.AreEqual(expected.Color, actual.Color);
        }

        /// <summary> 
        /// A Test for get Current layer
        /// </summary>
        [TestMethod()]
        public void CurrentlayerGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.Color = "ARGBColor:255:255:255:255";

            LayerMap_Accessor expected = new LayerMap_Accessor(layer);
            expected.MapType = LayerMapType.Local;
            expected.RangeDisplayName = "Sheet_1";
            target.currentLayer = expected;
            LayerMap_Accessor actual = target.Currentlayer;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A Test for set Current layer
        /// </summary>
        [TestMethod()]
        public void CurrentlayerSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.Color = "ARGBColor:255:255:255:255";

            LayerMap_Accessor expected = new LayerMap_Accessor(layer);
            expected.MapType = LayerMapType.Local;
            expected.HeaderRowData = new Collection<string>();
            expected.MappedColumnType = new Collection<ColumnType>();
            expected.RangeDisplayName = "Sheet_1";
            target.Currentlayer = expected;
            Assert.AreEqual(expected, target.Currentlayer);
        }

        /// <summary>
        /// A test for get DistanceUnits
        /// </summary>
        [TestMethod()]
        public void DistanceUnitsGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.distanceUnits = new Collection<KeyValuePair<AltUnit, string>>();
            target.distanceUnits.Add(new System.Collections.Generic.KeyValuePair<AltUnit, string>(AltUnit.Feet, "Feet"));
            ReadOnlyCollection<KeyValuePair<AltUnit, string>> actual = target.DistanceUnits;
            if (actual == null || actual.Count == 0)
            {
                Assert.Fail("Distance units could not be fetched.");
            }
            Assert.AreEqual(target.distanceUnits.Count, actual.Count);
        }

        /// <summary>
        /// A test for get EndDate
        /// </summary>
        [TestMethod()]
        public void EndDateGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.EndTime = DateTime.Now;

            LayerMap_Accessor expected = new LayerMap_Accessor(layer);
            expected.MapType = LayerMapType.Local;
            expected.HeaderRowData = new Collection<string>();
            expected.MappedColumnType = new Collection<ColumnType>();
            expected.RangeDisplayName = "Sheet_1";
            target.Currentlayer = expected;
            DateTime actual = target.EndDate;
            Assert.AreEqual(expected.LayerDetails.EndTime, actual);
        }

        /// <summary>
        /// A test  for set EndDate
        /// </summary>
        [TestMethod()]
        public void EndDateSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.EndTime = DateTime.Now;
            DateTime expected = DateTime.Now;
            LayerMap_Accessor layerMap = new LayerMap_Accessor(layer);
            layerMap.MapType = LayerMapType.Local;
            layerMap.HeaderRowData = new Collection<string>();
            layerMap.MappedColumnType = new Collection<ColumnType>();
            layerMap.RangeDisplayName = "Sheet_1";
            target.currentLayer = layerMap;
            target.BeginDate = expected;
            target.EndDate = expected;
            Assert.AreEqual(expected, target.EndDate);
        }

        /// <summary>
        /// A test  for set default EndDate
        /// </summary>
        [TestMethod()]
        public void EndDateDefaultSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            DateTime expected = DateTime.MaxValue;
            target.EndDate = DateTime.Now;
            Assert.AreEqual(expected, target.EndDate);
        }

        /// <summary>
        /// A test for get FadeTime
        /// </summary>
        [TestMethod()]
        public void FadeTimeGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.FadeSpan = new TimeSpan(10, 10, 10);

            LayerMap_Accessor layerMap = new LayerMap_Accessor(layer);
            layerMap.MapType = LayerMapType.Local;
            layerMap.HeaderRowData = new Collection<string>();
            layerMap.MappedColumnType = new Collection<ColumnType>();
            layerMap.RangeDisplayName = "Sheet_1";
            target.currentLayer = layerMap;

            string actual = target.FadeTime;
            Assert.AreEqual(layer.FadeSpan.ToString(), actual);
        }

        /// <summary>
        /// A test for get default FadeTime
        /// </summary>
        [TestMethod()]
        public void FadeTimeDefaultGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            string actual = target.FadeTime;
            Assert.AreEqual(string.Empty, actual);
        }

        /// <summary>
        /// A test for set FadeTime
        /// </summary>
        [TestMethod()]
        public void FadeTimeSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            TimeSpan expected = new TimeSpan(10, 10, 10);
            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.FadeSpan = expected;

            LayerMap_Accessor layerMap = new LayerMap_Accessor(layer);
            layerMap.MapType = LayerMapType.Local;
            layerMap.HeaderRowData = new Collection<string>();
            layerMap.MappedColumnType = new Collection<ColumnType>();
            layerMap.RangeDisplayName = "Sheet_1";
            target.currentLayer = layerMap;
            target.FadeTime = expected.ToString();
            TimeSpan span = new TimeSpan(0, 0, 0);
            if (string.IsNullOrEmpty(target.FadeTime) || !TimeSpan.TryParse(target.FadeTime, out span))
            {
                Assert.Fail("Time span not set.");
            }
            Assert.AreEqual(expected.ToString(), target.FadeTime);
        }

        /// <summary>
        /// A Test  for get LayerDataDisplayName
        /// </summary>
        [TestMethod()]
        public void LayerDataDisplayNameGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.layerDataDisplayName = "Layer";
            string actual = target.LayerDataDisplayName;
            if (string.IsNullOrEmpty(actual) || !actual.Equals("Layer"))
            {
                Assert.Fail("layer display name not fetched.");
            }
            Assert.AreEqual(target.layerDataDisplayName, actual);
        }

        /// <summary>
        /// A Test  for set LayerDataDisplayName
        /// </summary>
        [TestMethod()]
        public void LayerDataDisplayNameSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.LayerDataDisplayName = "Layer";
            if (string.IsNullOrEmpty(target.LayerDataDisplayName) || !target.LayerDataDisplayName.Equals("Layer"))
            {
                Assert.Fail("layer display name not set.");
            }
            Assert.AreEqual("Layer", target.layerDataDisplayName);
        }

        /// <summary>
        /// A test for get FadeTypes
        /// </summary>
        [TestMethod()]
        public void FadeTypesGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.fadetypes = new Collection<KeyValuePair<FadeType, string>>();
            target.fadetypes.Add(new KeyValuePair<FadeType, string>(FadeType.Both, "Both"));
            ReadOnlyCollection<KeyValuePair<FadeType, string>> actual = target.FadeTypes;
            if (actual == null || actual.Count == 0)
            {
                Assert.Fail("Fade type could not be fetched.");
            }
            Assert.AreEqual(target.fadetypes.Count, actual.Count);
        }

        /// <summary>
        /// A test  for get HoverTextColumnList
        /// </summary>
        [TestMethod()]
        public void HoverTextColumnListGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.hoverTextColumnList = new ObservableCollection<KeyValuePair<int, string>>();
            target.hoverTextColumnList.Add(new KeyValuePair<int, string>(1, "1"));
            ObservableCollection<KeyValuePair<int, string>> actual = target.HoverTextColumnList;
            if (actual == null || actual.Count == 0)
            {
                Assert.Fail("Hover text could not be fetched.");
            }
            Assert.AreEqual(target.hoverTextColumnList.Count, actual.Count);
        }

        /// <summary>
        /// A test  for set HoverTextColumnList
        /// </summary>
        [TestMethod()]
        public void HoverTextColumnListSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            ObservableCollection<KeyValuePair<int, string>> expected = new ObservableCollection<KeyValuePair<int, string>>();
            expected.Add(new KeyValuePair<int, string>(1, "1"));
            target.HoverTextColumnList = expected;
            if (target.HoverTextColumnList == null || target.HoverTextColumnList.Count == 0)
            {
                Assert.Fail("Hover text could not be set.");
            }
            Assert.AreEqual(expected.Count, target.HoverTextColumnList.Count);
        }

        /// <summary>
        /// A Test  for get LayerOpacity
        /// </summary>
        [TestMethod()]
        public void LayerOpacityGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Collection<double> ticks = new Collection<double>();
            ticks.Add(10);
            SliderViewModel slider = new SliderViewModel(ticks);
            slider.SelectedSliderValue = 10;
            target.layerOpacity = slider;
            SliderViewModel actual = target.LayerOpacity;
            Assert.AreEqual(10, actual.SelectedSliderValue);
        }

        /// <summary>
        /// A Test  for get Layers
        /// </summary>
        [TestMethod()]
        public void LayersGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.layers = new ObservableCollection<LayerMapDropDownViewModel>();
            target.layers.Add(new LayerMapDropDownViewModel() { ID = "1", Name = "1" });
            ObservableCollection<LayerMapDropDownViewModel> actual = target.Layers;
            if (actual == null || actual.Count == 0)
            {
                Assert.Fail("Layers could not be fetched.");
            }
            Assert.AreEqual(target.layers.Count, actual.Count);
        }

        /// <summary>
        /// A Test  for set Layers
        /// </summary>
        [TestMethod()]
        public void LayersSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            ObservableCollection<LayerMapDropDownViewModel> layers = new ObservableCollection<LayerMapDropDownViewModel>();
            layers.Add(new LayerMapDropDownViewModel() { ID = "1", Name = "1" });
            target.Layers = layers;
            if (target.Layers == null || target.Layers.Count == 0)
            {
                Assert.Fail("Layers could not be set.");
            }
            Assert.AreEqual(layers.Count, target.layers.Count);
        }

        /// <summary>
        /// A Test  for get ReferenceGroups
        /// </summary>
        [TestMethod()]
        public void ReferenceGroupsGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.referenceGroups = new ObservableCollection<GroupViewModel>();
            target.referenceGroups.Add(new GroupViewModel() { Name = "Layer" });
            ObservableCollection<GroupViewModel> actual = target.ReferenceGroups;
            if (actual == null || actual.Count == 0)
            {
                Assert.Fail("ReferenceGroups could not be fetched.");
            }
            Assert.AreEqual(target.referenceGroups.Count, actual.Count);
        }

        /// <summary>
        /// Test  for set ReferenceGroups
        /// </summary>
        [TestMethod()]
        public void ReferenceGroupsSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            ObservableCollection<GroupViewModel> referenceGroups = new ObservableCollection<GroupViewModel>();
            referenceGroups.Add(new GroupViewModel() { Name = "Layer" });
            target.ReferenceGroups = referenceGroups;
            if (target.ReferenceGroups == null || target.ReferenceGroups.Count == 0)
            {
                Assert.Fail("Reference Groups could not be set.");
            }
            Assert.AreEqual(referenceGroups.Count, target.ReferenceGroups.Count);
        }

        /// <summary>
        /// A Test  for get RightAscentionUnits
        /// </summary>
        [TestMethod()]
        public void RightAscentionUnitsGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.rightAscentionUnits = new Collection<KeyValuePair<AngleUnit, string>>();
            target.rightAscentionUnits.Add(new KeyValuePair<AngleUnit, string>(AngleUnit.Hours, "Hours"));
            ReadOnlyCollection<KeyValuePair<AngleUnit, string>> actual = target.RightAscentionUnits;
            if (actual == null || actual.Count == 0)
            {
                Assert.Fail("RightAscentionUnits could not be fetched.");
            }
            Assert.AreEqual(target.rightAscentionUnits.Count, actual.Count);
        }

        /// <summary>
        /// A Test  for get ScaleFactor
        /// </summary>
        [TestMethod()]
        public void ScaleFactorGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Collection<double> ticks = new Collection<double>();
            ticks.Add(1);
            ticks.Add(2);
            SliderViewModel scaleFactor = new SliderViewModel(ticks);
            target.scaleFactor = scaleFactor;
            SliderViewModel actual = scaleFactor;
            if (actual == null || actual.SliderTicks.Count == 0)
            {
                Assert.Fail("ScaleFactor could not be fetched.");
            }
            Assert.AreEqual(target.scaleFactor.SliderTicks.Count, actual.SliderTicks.Count);
        }

        /// <summary>
        /// A Test  for get ScaleRelatives
        /// </summary>
        [TestMethod()]
        public void ScaleRelativesGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.scaleRelatives = new Collection<KeyValuePair<ScaleRelativeType, string>>();
            target.scaleRelatives.Add(new KeyValuePair<ScaleRelativeType, string>(ScaleRelativeType.Screen, "Screen"));
            ReadOnlyCollection<KeyValuePair<ScaleRelativeType, string>> actual = target.ScaleRelatives;
            if (actual == null || actual.Count == 0)
            {
                Assert.Fail("ScaleRelatives could not be fetched.");
            }
            Assert.AreEqual(target.scaleRelatives.Count, actual.Count);
        }

        /// <summary>
        /// Test  for get ScaleTypes
        /// </summary>
        [TestMethod()]
        public void ScaleTypesGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.scaleTypes = new Collection<KeyValuePair<ScaleType, string>>();
            target.scaleTypes.Add(new KeyValuePair<ScaleType, string>(ScaleType.Linear, "linear"));
            ReadOnlyCollection<KeyValuePair<ScaleType, string>> actual = target.ScaleTypes;
            if (actual == null || actual.Count == 0)
            {
                Assert.Fail("ScaleTypes could not be fetched.");
            }
            Assert.AreEqual(target.scaleTypes.Count, actual.Count);
        }

        /// <summary>
        /// A Test  for get SelectedDistanceUnit
        /// </summary>
        [TestMethod()]
        public void SelectedDistanceUnitGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.selectedDistanceUnit = new KeyValuePair<AltUnit, string>(AltUnit.Feet, "Feet");
            KeyValuePair<AltUnit, string> actual = target.SelectedDistanceUnit;
            if (!actual.Key.Equals(AltUnit.Feet))
            {
                Assert.Fail("SelectedDistanceUnit could not be fetched.");
            }
            Assert.AreEqual(target.selectedDistanceUnit.Key, AltUnit.Feet);
        }

        /// <summary>
        /// A Test  for set SelectedDistanceUnit
        /// </summary>
        [TestMethod()]
        public void SelectedDistanceUnitSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            KeyValuePair<AltUnit, string> selectedDistanceUnit = new KeyValuePair<AltUnit, string>(AltUnit.Feet, "Feet");
            target.SelectedDistanceUnit = selectedDistanceUnit;
            if (!target.SelectedDistanceUnit.Key.Equals(AltUnit.Feet))
            {
                Assert.Fail("SelectedDistanceUnit could not be set.");
            }
            Assert.AreEqual(target.SelectedDistanceUnit.Key, AltUnit.Feet);
        }

        /// <summary>
        /// A Test  for get SelectedFadeType
        /// </summary>
        [TestMethod()]
        public void SelectedFadeTypeGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.selectedFadeType = new KeyValuePair<FadeType, string>(FadeType.Both, "Both");
            KeyValuePair<FadeType, string> actual = target.SelectedFadeType;
            if (!actual.Key.Equals(FadeType.Both))
            {
                Assert.Fail("SelectedFadeType could not be fetched.");
            }
            Assert.AreEqual(actual.Key, FadeType.Both);
        }

        /// <summary>
        /// A Test  for set SelectedFadeType
        /// </summary>
        [TestMethod()]
        public void SelectedFadeTypeSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            KeyValuePair<FadeType, string> selectedFadeType = new KeyValuePair<FadeType, string>(FadeType.Both, "Both");
            target.SelectedFadeType = selectedFadeType;
            if (!target.SelectedFadeType.Key.Equals(FadeType.Both))
            {
                Assert.Fail("SelectedFadeType could not be fetched.");
            }
            Assert.AreEqual(target.SelectedFadeType.Key, FadeType.Both);
        }

        /// <summary>
        /// A Test  for get SelectedGroup
        /// </summary>
        [TestMethod()]
        public void SelectedGroupGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.selectedGroup = new Group("Sun", GroupType.LayerGroup, null);
            Group actual = target.SelectedGroup;
            if (actual == null || !actual.GroupType.Equals(GroupType.LayerGroup))
            {
                Assert.Fail("SelectedGroup could not be fetched.");
            }
            Assert.AreEqual(actual.GroupType, GroupType.LayerGroup);
        }

        /// <summary>
        /// A Test  for set SelectedGroup
        /// </summary>
        [TestMethod()]
        public void SelectedGroupSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Group selectedGroup = new Group("Sun", GroupType.LayerGroup, null);
            target.SelectedGroup = selectedGroup;
            if (target.SelectedGroup == null || !target.SelectedGroup.GroupType.Equals(GroupType.LayerGroup))
            {
                Assert.Fail("SelectedGroup could not be set.");
            }
            Assert.AreEqual(target.SelectedGroup.GroupType, GroupType.LayerGroup);
        }

        /// <summary>
        /// A Test  for get SelectedGroupText
        /// </summary>
        [TestMethod()]
        public void SelectedGroupTextGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.selectedGroupText = "group";
            string actual = target.SelectedGroupText;
            if (string.IsNullOrEmpty(actual))
            {
                Assert.Fail("SelectedGroupText could not be fetched.");
            }
            Assert.AreEqual(target.selectedGroupText, actual);
        }

        /// <summary>
        /// A Test  for set_SelectedGroupText
        /// </summary>
        [TestMethod()]
        public void SelectedGroupTextSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            string selectedGroupText = "group";
            target.SelectedGroupText = selectedGroupText;
            if (string.IsNullOrEmpty(target.SelectedGroupText))
            {
                Assert.Fail("SelectedGroupText could not be set.");
            }
            Assert.AreEqual(selectedGroupText, target.SelectedGroupText);
        }

        /// <summary>
        /// A Test  for get SelectedHoverText
        /// </summary>
        [TestMethod()]
        public void SelectedHoverTextGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.selectedHoverText = new KeyValuePair<int, string>(1, "1");
            KeyValuePair<int, string> actual = target.SelectedHoverText;
            if (!actual.Key.Equals(1))
            {
                Assert.Fail("SelectedHoverText could not be fetched.");
            }
            Assert.AreEqual(actual.Key, 1);
        }

        /// <summary>
        /// A Test  for set SelectedHoverText
        /// </summary>
        [TestMethod()]
        public void SelectedHoverTextSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            KeyValuePair<int, string> selectedHoverText = new KeyValuePair<int, string>(1, "1");
            target.SelectedHoverText = selectedHoverText;
            if (!target.SelectedHoverText.Key.Equals(1))
            {
                Assert.Fail("SelectedHoverText could not be set.");
            }
            Assert.AreEqual(target.SelectedHoverText.Key, 1);
        }

        /// <summary>
        /// A Test  for get_SelectedLayerMapDropDown
        /// </summary>
        [TestMethod()]
        public void SelectedLayerMapDropDownGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.selectedLayerMapDropDown = new LayerMapDropDownViewModel();
            target.selectedLayerMapDropDown.ID = "1";
            target.selectedLayerMapDropDown.Name = "1";
            LayerMapDropDownViewModel actual = target.SelectedLayerMapDropDown;
            if (actual == null || !actual.Name.Equals("1"))
            {
                Assert.Fail("SelectedLayerMapDropDown could not be fetched.");
            }
            Assert.AreEqual("1", actual.Name);
        }

        /// <summary>
        /// A Test  for set SelectedLayerMapDropDown
        /// </summary>
        [TestMethod()]
        public void SelectedLayerMapDropDownSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            LayerMapDropDownViewModel selectedLayerMapDropDown = new LayerMapDropDownViewModel();
            selectedLayerMapDropDown.ID = "1";
            selectedLayerMapDropDown.Name = "1";
            target.SelectedLayerMapDropDown = selectedLayerMapDropDown;
            if (target.SelectedLayerMapDropDown == null || !target.SelectedLayerMapDropDown.Name.Equals("1"))
            {
                Assert.Fail("SelectedLayerMapDropDown could not be set.");
            }
            Assert.AreEqual("1", target.SelectedLayerMapDropDown.Name);
        }

        /// <summary>
        /// A Test  for get SelectedLayerName
        /// </summary>
        [TestMethod()]
        public void SelectedLayerNameGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.selectedLayerName = "Layer";
            string actual = target.SelectedLayerName;
            if (string.IsNullOrEmpty(actual))
            {
                Assert.Fail("SelectedLayerName could not be fetched.");
            }
            Assert.AreEqual(target.selectedLayerName, actual);
        }

        /// <summary>
        /// A Test  for set SelectedLayerName
        /// </summary>
        [TestMethod()]
        public void SelectedLayerNameSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            string selectedLayerName = "Layer";
            Layer layer = new Layer();
            layer.Name = "Layer1";

            LayerMap_Accessor layerMap = new LayerMap_Accessor(layer);
            layerMap.MapType = LayerMapType.Local;
            layerMap.HeaderRowData = new Collection<string>();
            layerMap.MappedColumnType = new Collection<ColumnType>();
            layerMap.RangeDisplayName = "Sheet_1";
            target.currentLayer = layerMap;
            target.SelectedLayerName = selectedLayerName;
            if (string.IsNullOrEmpty(target.SelectedLayerName))
            {
                Assert.Fail("SelectedLayerName could not be set.");
            }
            Assert.AreEqual(selectedLayerName, target.SelectedLayerName);
        }

        /// <summary>
        /// A test  for get IsCallOutVisible
        /// </summary>
        [TestMethod()]
        public void IsCallOutVisibleGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.isCallOutVisible = true;
            bool actual = target.IsCallOutVisible;
            Assert.IsTrue(actual);
        }

        /// <summary>
        /// A test  for set IsCallOutVisible
        /// </summary>
        [TestMethod()]
        public void IsCallOutVisibleSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.IsCallOutVisible = true;
            Assert.IsTrue(target.IsCallOutVisible);
        }

        /// <summary>
        /// A test  for get IsDeleteMappingEnabled
        /// </summary>
        [TestMethod()]
        public void IsDeleteMappingEnabledGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.isDeleteMappingEnabled = true;
            bool actual = target.IsDeleteMappingEnabled;
            Assert.IsTrue(actual);
        }

        /// <summary>
        /// A test  for set IsDeleteMappingEnabled
        /// </summary>
        [TestMethod()]
        public void IsDeleteMappingEnabledSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.IsDeleteMappingEnabled = true;
            Assert.IsTrue(target.IsDeleteMappingEnabled);
        }

        /// <summary>
        /// A Test  for get IsDistanceVisible
        /// </summary>
        [TestMethod()]
        public void IsDistanceVisibleGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.isDistanceVisible = true;
            bool actual = target.IsDistanceVisible;
            Assert.IsTrue(actual);
        }

        /// <summary>
        /// A test  for set IsDistanceVisible
        /// </summary>
        [TestMethod()]
        public void IsDistanceVisibleSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.IsDistanceVisible = true;
            Assert.IsTrue(target.IsDistanceVisible);
        }

        /// <summary>
        /// A test  for get IsGetLayerDataEnabled
        /// </summary>
        [TestMethod()]
        public void IsGetLayerDataEnabledGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.isGetLayerDataEnabled = true;
            bool actual = target.IsGetLayerDataEnabled;
            Assert.IsTrue(actual);
        }

        /// <summary>
        /// A test for set IsGetLayerDataEnabled
        /// </summary>
        [TestMethod()]
        public void IsGetLayerDataEnabledSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.IsGetLayerDataEnabled = true;
            Assert.IsTrue(target.IsGetLayerDataEnabled);
        }

        /// <summary>
        /// A test for get IsHelpTextVisible
        /// </summary>
        [TestMethod()]
        public void IsHelpTextVisibleGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.isHelpTextVisible = true;
            bool actual = target.IsHelpTextVisible;
            Assert.IsTrue(actual);
        }

        /// <summary>
        /// A Test for IsHelpTextVisible
        /// </summary>
        [TestMethod()]
        public void IsHelpTextVisibleSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.IsHelpTextVisible = true;
            Assert.IsTrue(target.IsHelpTextVisible);
        }

        /// <summary>
        /// A Test  for get IsMarkerTabEnabled
        /// </summary>
        [TestMethod()]
        public void IsMarkerTabEnabledGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.isMarkerTabEnabled = true;
            bool actual = target.IsMarkerTabEnabled;
            Assert.IsTrue(actual);
        }

        /// <summary>
        /// A test  for set IsMarkerTabEnabled
        /// </summary>
        [TestMethod()]
        public void IsMarkerTabEnabledSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.IsMarkerTabEnabled = true;
            Assert.IsTrue(target.IsMarkerTabEnabled);
        }

        /// <summary>
        /// A Test  for get IsRAUnitVisible
        /// </summary>
        [TestMethod()]
        public void IsRAUnitVisibleGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.isRAUnitVisible = true;
            bool actual = target.IsRAUnitVisible;
            Assert.IsTrue(actual);
        }

        /// <summary>
        /// A Test  for set IsRAUnitVisible
        /// </summary>
        [TestMethod()]
        public void IsRAUnitVisibleSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.IsRAUnitVisible = true;
            Assert.IsTrue(target.IsRAUnitVisible);
        }

        /// <summary>
        /// A Test  for get IsReferenceGroupEnabled
        /// </summary>
        [TestMethod()]
        public void IsReferenceGroupEnabledGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.isReferenceGroupEnabled = true;
            bool actual = target.IsReferenceGroupEnabled;
            Assert.IsTrue(actual);
        }

        /// <summary>
        /// A Test  for set IsReferenceGroupEnabled
        /// </summary>
        [TestMethod()]
        public void IsReferenceGroupEnabledSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.IsReferenceGroupEnabled = true;
            Assert.IsTrue(target.IsReferenceGroupEnabled);
        }

        /// <summary>
        /// A Test  for get IsShowRangeEnabled
        /// </summary>
        [TestMethod()]
        public void IsShowRangeEnabledGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.isShowRangeEnabled = true;
            bool actual = target.IsShowRangeEnabled;
            Assert.IsTrue(actual);
        }

        /// <summary>
        /// A Test  for set_IsShowRangeEnabled
        /// </summary>
        [TestMethod()]
        public void IsShowRangeEnabledSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.IsShowRangeEnabled = true;
            Assert.IsTrue(target.IsShowRangeEnabled);
        }

        /// <summary>
        /// A Test  for get IsTabVisible
        /// </summary>
        [TestMethod()]
        public void IsTabVisibleGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.isTabVisible = true;
            bool actual = target.IsTabVisible;
            Assert.IsTrue(actual);
        }

        /// <summary>
        /// A Test  for set IsTabVisible
        /// </summary>
        [TestMethod()]
        public void IsTabVisibleSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.IsTabVisible = true;
            Assert.IsTrue(target.IsTabVisible);
        }

        /// <summary>
        /// A Test  for get IsUpdateLayerEnabled
        /// </summary>
        [TestMethod]
        public void IsUpdateLayerEnabledGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.isUpdateLayerEnabled = true;
            bool actual = target.IsUpdateLayerEnabled;
            Assert.IsTrue(actual);
        }

        /// <summary>
        /// A Test  for set IsUpdateLayerEnabled
        /// </summary>
        [TestMethod()]
        public void IsUpdateLayerEnabledSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.IsUpdateLayerEnabled = true;
            Assert.IsTrue(target.IsUpdateLayerEnabled);
        }

        /// <summary>
        /// A Test  for get isViewInWWTEnable
        /// </summary>
        [TestMethod()]
        public void IsViewInWWTEnableGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.isViewInWWTEnabled = true;
            bool actual = target.isViewInWWTEnabled;
            Assert.IsTrue(actual);
        }

        /// <summary>
        /// A Test  for set isViewInWWTEnable
        /// </summary>
        [TestMethod()]
        public void IsViewInWWTEnableSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.isViewInWWTEnabled = true;
            Assert.IsTrue(target.isViewInWWTEnabled);
        }

        /// <summary>
        /// A test for GetLayerNameOnMapType for local layer
        /// </summary>
        [TestMethod()]
        public void GetLayerNameOnMapTypeLocalTest()
        {
            Layer layer = new Layer();
            layer.Name = "Layer1";

            LayerMap layerMap = new LayerMap(layer);
            layerMap.MapType = LayerMapType.Local;
            layerMap.HeaderRowData = new Collection<string>();
            layerMap.MappedColumnType = new Collection<ColumnType>();
            string actual = LayerDetailsViewModel.GetLayerNameOnMapType(layerMap, layer.Name);
            if (string.IsNullOrEmpty(actual) || !actual.Contains(Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.LayerLocalText))
            {
                Assert.Fail("GetLayerNameOnMapType for local layer failed.");
            }
        }

        /// <summary>
        /// A test for GetLayerNameOnMapType local in WWT, not in sync
        /// </summary>
        [TestMethod()]
        public void GetLayerNameOnMapTypeLocalInWWTNotSynchTest()
        {
            Layer layer = new Layer();
            layer.Name = "Layer1";

            LayerMap layerMap = new LayerMap(layer);
            layerMap.MapType = LayerMapType.LocalInWWT;
            layerMap.HeaderRowData = new Collection<string>();
            layerMap.MappedColumnType = new Collection<ColumnType>();
            layerMap.IsNotInSync = true;
            string actual = LayerDetailsViewModel.GetLayerNameOnMapType(layerMap, layer.Name);
            if (string.IsNullOrEmpty(actual) || !actual.Contains(Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.LayerLocalText))
            {
                Assert.Fail("GetLayerNameOnMapType for local in WWT layer failed.");
            }
        }

        /// <summary>
        /// A test for GetLayerNameOnMapType local in WWT, in sync
        /// </summary>
        [TestMethod()]
        public void GetLayerNameOnMapTypeLocalInWWTSynchTest()
        {
            Layer layer = new Layer();
            layer.Name = "Layer1";

            LayerMap layerMap = new LayerMap(layer);
            layerMap.MapType = LayerMapType.LocalInWWT;
            layerMap.HeaderRowData = new Collection<string>();
            layerMap.MappedColumnType = new Collection<ColumnType>();
            layerMap.IsNotInSync = false;
            string actual = LayerDetailsViewModel.GetLayerNameOnMapType(layerMap, layer.Name);
            if (string.IsNullOrEmpty(actual) || !actual.Contains(Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.LayerLocalInWWTText))
            {
                Assert.Fail("GetLayerNameOnMapType for local in WWT layer failed.");
            }
        }

        /// <summary>
        /// A test for GetLayerNameOnMapType WWT, in sync
        /// </summary>
        [TestMethod()]
        public void GetLayerNameOnMapTypeWWTTest()
        {
            Layer layer = new Layer();
            layer.Name = "Layer1";

            LayerMap layerMap = new LayerMap(layer);
            layerMap.MapType = LayerMapType.WWT;
            layerMap.HeaderRowData = new Collection<string>();
            layerMap.MappedColumnType = new Collection<ColumnType>();
            string actual = LayerDetailsViewModel.GetLayerNameOnMapType(layerMap, layer.Name);
            if (string.IsNullOrEmpty(actual) || !actual.Equals(layer.Name))
            {
                Assert.Fail("GetLayerNameOnMapType for local in WWT layer failed.");
            }
        }

        /// <summary>
        /// A Test for get SelectedLayerText
        /// </summary>
        [TestMethod()]
        public void SelectedLayerTextGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            string expected = "SelectedLayer";
            target.selectedLayerText = expected;
            string actual = target.SelectedLayerText;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A Test for set SelectedLayerText
        /// </summary>
        [TestMethod()]
        public void SelectedLayerTextSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            string expected = "SelectedLayer";
            target.SelectedLayerText = expected;
            Assert.AreEqual(expected, target.SelectedLayerText);
        }

        /// <summary>
        /// A Test for get SelectedRAUnit
        /// </summary>
        [TestMethod()]
        public void SelectedRAUnitGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            KeyValuePair<AngleUnit, string> expected = new KeyValuePair<AngleUnit, string>(AngleUnit.Degrees, "Degrees");
            target.selectedRAUnit = expected;
            KeyValuePair<AngleUnit, string> actual = target.SelectedRAUnit;
            Assert.AreEqual(expected.Key, actual.Key);
        }

        /// <summary>
        /// A Test for set SelectedRAUnit
        /// </summary>
        [TestMethod()]
        public void SelectedRAUnitSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            KeyValuePair<AngleUnit, string> expected = new KeyValuePair<AngleUnit, string>(AngleUnit.Degrees, "Degrees");
            target.SelectedRAUnit = expected;
            Assert.AreEqual(expected.Key, target.SelectedRAUnit.Key);
        }

        /// <summary>
        /// A Test for get SelectedScaleFactor
        /// </summary>
        [TestMethod()]
        public void SelectedScaleFactorGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            double expected = 1;
            target.scaleFactor.SelectedSliderValue = expected;
            double actual = target.scaleFactor.SelectedSliderValue;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A Test for set SelectedScaleFactor
        /// </summary>
        [TestMethod()]
        public void SelectedScaleFactorSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            double expected = 1;
            target.scaleFactor.SelectedSliderValue = expected;
            Assert.AreEqual(expected, target.scaleFactor.SelectedSliderValue);
        }

        /// <summary>
        /// A Test for get SelectedScaleRelative
        /// </summary>
        [TestMethod()]
        public void SelectedScaleRelativeGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            KeyValuePair<ScaleRelativeType, string> expected = new KeyValuePair<ScaleRelativeType, string>(ScaleRelativeType.Screen, "Screen");
            target.selectedScaleRelative = expected;
            KeyValuePair<ScaleRelativeType, string> actual = target.SelectedScaleRelative;
            Assert.AreEqual(expected.Key, actual.Key);
        }

        /// <summary>
        /// A Test for set SelectedScaleRelative
        /// </summary>
        [TestMethod()]
        public void SelectedScaleRelativeSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            KeyValuePair<ScaleRelativeType, string> expected = new KeyValuePair<ScaleRelativeType, string>(ScaleRelativeType.Screen, "Screen");
            target.SelectedScaleRelative = expected;
            Assert.AreEqual(expected.Key, target.SelectedScaleRelative.Key);
        }

        /// <summary>
        /// A Test for get SelectedScaleType
        /// </summary>
        [TestMethod()]
        public void SelectedScaleTypeGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            KeyValuePair<ScaleType, string> expected = new KeyValuePair<ScaleType, string>(ScaleType.Constant, "Const");
            target.selectedScaleType = expected;
            KeyValuePair<ScaleType, string> actual = target.SelectedScaleType;
            Assert.AreEqual(expected.Key, actual.Key);
        }

        /// <summary>
        /// A Test for set SelectedScaleType
        /// </summary>
        [TestMethod()]
        public void SelectedScaleTypeSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            KeyValuePair<ScaleType, string> expected = new KeyValuePair<ScaleType, string>(ScaleType.Constant, "Const");
            target.SelectedScaleType = expected;
            Assert.AreEqual(expected.Key, target.SelectedScaleType.Key);
        }

        /// <summary>
        /// A Test for get SelectedSize
        /// </summary>
        [TestMethod()]
        public void SelectedSizeGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            KeyValuePair<int, string> expected = new KeyValuePair<int, string>(1, "1");
            target.selectedSizeColumn = expected;
            KeyValuePair<int, string> actual = target.SelectedSize;
            Assert.AreEqual(expected.Key, actual.Key);
        }

        /// <summary>
        /// A Test stub for set SelectedSize
        /// </summary>
        [TestMethod()]
        public void SelectedSizeSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            KeyValuePair<int, string> expected = new KeyValuePair<int, string>(1, "1");
            target.SelectedSize = expected;
            Assert.AreEqual(expected.Key, target.SelectedSize.Key);
        }

        /// <summary>
        /// A Test for get SelectedTabIndex
        /// </summary>
        [TestMethod()]
        public void SelectedTabIndexGetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            int expected = 1;
            target.selectedTabIndex = expected;
            int actual = target.SelectedTabIndex;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A Test for set SelectedTabIndex
        /// </summary>
        [TestMethod()]
        public void SelectedTabIndexSetTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            int expected = 1;
            target.SelectedTabIndex = expected;
            Assert.AreEqual(expected, target.SelectedTabIndex);
        }

        /// <summary>
        /// A Test for SetDistanceUnitVisibility for a depth column 
        /// </summary>
        [TestMethod()]
        public void SetDistanceUnitVisibilityDepthColumnTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            bool isDepthColumnSelected = true;
            ObservableCollection<ColumnViewModel> expected = new ObservableCollection<ColumnViewModel>();
            expected.Add(new ColumnViewModel() { ExcelHeaderColumn = "Depth", SelectedWWTColumn = new Column(ColumnType.Depth, "Depth", new Collection<string>() { "Depth" }) });
            target.ColumnsView = expected;
            target.SetDistanceUnitVisibility(isDepthColumnSelected);
            Assert.IsTrue(target.IsDistanceVisible);
            Assert.AreEqual(AltUnit.Kilometers, target.SelectedDistanceUnit.Key);
        }

        /// <summary>
        /// A Test for SetDistanceUnitVisibility for a non-depth column
        /// </summary>
        [TestMethod()]
        public void SetDistanceUnitVisibilityNotDepthColumnTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            bool isDepthColumnSelected = false;
            ObservableCollection<ColumnViewModel> expected = new ObservableCollection<ColumnViewModel>();
            expected.Add(new ColumnViewModel() { ExcelHeaderColumn = "DEC", SelectedWWTColumn = new Column(ColumnType.Dec, "DEC", new Collection<string>() { "DEC" }) });
            target.ColumnsView = expected;
            target.SetDistanceUnitVisibility(isDepthColumnSelected);
            Assert.IsFalse(target.IsDistanceVisible);
        }

        /// <summary>
        /// A Test for SetMarkerTabVisibility for geo column
        /// </summary>
        [TestMethod()]
        public void SetMarkerTabVisibilityGeoTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            ObservableCollection<ColumnViewModel> expected = new ObservableCollection<ColumnViewModel>();
            expected.Add(new ColumnViewModel() { ExcelHeaderColumn = "Geo", SelectedWWTColumn = new Column(ColumnType.Geo, "Geo", new Collection<string>() { "Geo" }) });
            target.ColumnsView = expected;
            target.SetMarkerTabVisibility();
            Assert.IsTrue(target.IsMarkerTabEnabled);
        }

        /// <summary>
        /// A Test for SetMarkerTabVisibility for geo column
        /// </summary>
        [TestMethod()]
        public void SetMarkerTabVisibilityNonGeoTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            ObservableCollection<ColumnViewModel> expected = new ObservableCollection<ColumnViewModel>();
            expected.Add(new ColumnViewModel() { ExcelHeaderColumn = "Dec", SelectedWWTColumn = new Column(ColumnType.Dec, "Dec", new Collection<string>() { "Dec" }) });
            target.ColumnsView = expected;
            target.SetMarkerTabVisibility();
            Assert.IsTrue(target.IsMarkerTabEnabled);
        }

        /// <summary>
        /// A Test for SetRAUnitVisibility for RA column
        /// </summary>
        [TestMethod()]
        public void SetRAUnitVisibilityRATest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            ObservableCollection<ColumnViewModel> expected = new ObservableCollection<ColumnViewModel>();
            expected.Add(new ColumnViewModel() { ExcelHeaderColumn = "RA", SelectedWWTColumn = new Column(ColumnType.RA, "RA", new Collection<string>() { "RA" }) });
            target.ColumnsView = expected;
            Layer layer = new Layer();
            layer.Name = "Layer1";
            layer.StartTime = DateTime.Now;
            layer.RAUnit = AngleUnit.Hours;

            LayerMap_Accessor localLayerMap = new LayerMap_Accessor(layer);
            localLayerMap.MapType = LayerMapType.Local;
            localLayerMap.RangeDisplayName = "Sheet_1";

            target.currentLayer = localLayerMap;
            target.SetRAUnitVisibility();
            Assert.IsTrue(target.IsRAUnitVisible);
            Assert.AreEqual(layer.RAUnit, target.SelectedRAUnit.Key);
        }

        /// <summary>
        /// A Test for SetRAUnitVisibility for non RA column
        /// </summary>
        [TestMethod()]
        public void SetRAUnitVisibilityNonRATest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            ObservableCollection<ColumnViewModel> expected = new ObservableCollection<ColumnViewModel>();
            expected.Add(new ColumnViewModel() { ExcelHeaderColumn = "Depth", SelectedWWTColumn = new Column(ColumnType.Depth, "Depth", new Collection<string>() { "Depth" }) });
            target.ColumnsView = expected;
            target.SetRAUnitVisibility();
            Assert.IsFalse(target.IsRAUnitVisible);
            Assert.AreNotEqual(ColumnType.RA, target.SelectedRAUnit.Key);
        }

        /// <summary>
        /// A test for BuildLayerOpacity
        /// </summary>
        [TestMethod()]
        public void BuildLayerOpacityTest()
        {
            Collection<double> expected = new Collection<double>();
            for (int i = 0; i <= 100; i++)
            {
                expected.Add(i);
            }
            Collection<double> actual;
            actual = LayerDetailsViewModel_Accessor.BuildLayerOpacity();
            Assert.AreEqual(expected.Count, actual.Count);
            int index = 0;
            foreach (double value in actual)
            {
                Assert.AreEqual(expected[index], value);
                index++;
            }
        }

        /// <summary>
        /// A test for PopulateDistanceUnits
        /// </summary>
        [TestMethod()]
        public void PopulateDistanceUnitsTest()
        {
            Collection<KeyValuePair<AltUnit, string>> expected = new Collection<KeyValuePair<AltUnit, string>>();
            expected.Add(new KeyValuePair<AltUnit, string>(AltUnit.Inches, Resources.DistanceInches));
            expected.Add(new KeyValuePair<AltUnit, string>(AltUnit.Feet, Resources.DistanceFeet));
            expected.Add(new KeyValuePair<AltUnit, string>(AltUnit.Miles, Resources.DistanceMiles));
            expected.Add(new KeyValuePair<AltUnit, string>(AltUnit.Meters, Resources.DistanceMeters));
            expected.Add(new KeyValuePair<AltUnit, string>(AltUnit.Kilometers, Resources.DistanceKiloMeters));
            expected.Add(new KeyValuePair<AltUnit, string>(AltUnit.AstronomicalUnits, Resources.DistanceAstronomicalUnits));
            expected.Add(new KeyValuePair<AltUnit, string>(AltUnit.LightYears, Resources.DistanceLightYears));
            expected.Add(new KeyValuePair<AltUnit, string>(AltUnit.Parsecs, Resources.DistanceParsecs));
            expected.Add(new KeyValuePair<AltUnit, string>(AltUnit.MegaParsecs, Resources.DistanceMegaParsecs));

            Collection<KeyValuePair<AltUnit, string>> actual;
            actual = LayerDetailsViewModel_Accessor.PopulateDistanceUnits();
            Assert.AreEqual(expected.Count, actual.Count);
            int index = 0;
            foreach (KeyValuePair<AltUnit, string> value in actual)
            {
                Assert.AreEqual(expected[index].Key, value.Key);
                Assert.AreEqual(expected[index].Value, value.Value);
                index++;
            }
        }

        /// <summary>
        /// A test for PopulateFadeType
        /// </summary>
        [TestMethod()]
        public void PopulateFadeTypeTest()
        {
            Collection<KeyValuePair<FadeType, string>> expected = new Collection<KeyValuePair<FadeType, string>>();
            expected.Add(new KeyValuePair<FadeType, string>(FadeType.None, Resources.FadeNone));
            expected.Add(new KeyValuePair<FadeType, string>(FadeType.In, Resources.FadeIn));
            expected.Add(new KeyValuePair<FadeType, string>(FadeType.Out, Resources.FadeOut));
            expected.Add(new KeyValuePair<FadeType, string>(FadeType.Both, Resources.FadeBoth));

            Collection<KeyValuePair<FadeType, string>> actual;
            actual = LayerDetailsViewModel_Accessor.PopulateFadeType();
            Assert.AreEqual(expected.Count, actual.Count);
            int index = 0;
            foreach (KeyValuePair<FadeType, string> value in actual)
            {
                Assert.AreEqual(expected[index].Key, value.Key);
                Assert.AreEqual(expected[index].Value, value.Value);
                index++;
            }
        }

        /// <summary>
        /// A test for PopulateRAUnits
        /// </summary>
        [TestMethod()]
        public void PopulateRAUnitsTest()
        {
            Collection<KeyValuePair<AngleUnit, string>> expected = new Collection<KeyValuePair<AngleUnit, string>>();
            expected.Add(new KeyValuePair<AngleUnit, string>(AngleUnit.Hours, Resources.RAHour));
            expected.Add(new KeyValuePair<AngleUnit, string>(AngleUnit.Degrees, Resources.RADegree));
            Collection<KeyValuePair<AngleUnit, string>> actual;
            actual = LayerDetailsViewModel_Accessor.PopulateRAUnits();
            Assert.AreEqual(expected.Count, actual.Count);
            int index = 0;
            foreach (KeyValuePair<AngleUnit, string> value in actual)
            {
                Assert.AreEqual(expected[index].Key, value.Key);
                Assert.AreEqual(expected[index].Value, value.Value);
                index++;
            }
        }

        /// <summary>
        /// A test for PopulateScaleRelatives
        /// </summary>
        [TestMethod()]
        public void PopulateScaleRelativesTest()
        {
            Collection<KeyValuePair<ScaleRelativeType, string>> expected = new Collection<KeyValuePair<ScaleRelativeType, string>>();
            expected.Add(new KeyValuePair<ScaleRelativeType, string>(ScaleRelativeType.World, Resources.ScaleRelativeWorld));
            expected.Add(new KeyValuePair<ScaleRelativeType, string>(ScaleRelativeType.Screen, Resources.ScaleRelativeScreen));

            Collection<KeyValuePair<ScaleRelativeType, string>> actual;
            actual = LayerDetailsViewModel_Accessor.PopulateScaleRelatives();
            Assert.AreEqual(expected.Count, actual.Count);
            int index = 0;
            foreach (KeyValuePair<ScaleRelativeType, string> value in actual)
            {
                Assert.AreEqual(expected[index].Key, value.Key);
                Assert.AreEqual(expected[index].Value, value.Value);
                index++;
            }
        }

        /// <summary>
        /// A test for PopulateScaleType
        /// </summary>
        [TestMethod()]
        public void PopulateScaleTypeTest()
        {
            Collection<KeyValuePair<ScaleType, string>> expected = new Collection<KeyValuePair<ScaleType, string>>();
            expected.Add(new KeyValuePair<ScaleType, string>(ScaleType.Power, Resources.ScaleTypePower));
            expected.Add(new KeyValuePair<ScaleType, string>(ScaleType.Constant, Resources.ScaleTypeConstant));
            expected.Add(new KeyValuePair<ScaleType, string>(ScaleType.Linear, Resources.ScaleTypeLinear));
            expected.Add(new KeyValuePair<ScaleType, string>(ScaleType.Log, Resources.ScaleTypeLog));
            expected.Add(new KeyValuePair<ScaleType, string>(ScaleType.StellarMagnitude, Resources.ScaleTypeStellarMagnitude));

            Collection<KeyValuePair<ScaleType, string>> actual;
            actual = LayerDetailsViewModel_Accessor.PopulateScaleType();
            Assert.AreEqual(expected.Count, actual.Count);
            int index = 0;
            foreach (KeyValuePair<ScaleType, string> value in actual)
            {
                Assert.AreEqual(expected[index].Key, value.Key);
                Assert.AreEqual(expected[index].Value, value.Value);
                index++;
            }
        }

        /// <summary>
        /// A test for SetDefaultBackground
        /// </summary>
        [TestMethod()]
        public void SetDefaultBackgroundTest()
        {
            SolidColorBrush expected = new SolidColorBrush(System.Windows.Media.Color.FromArgb(System.Drawing.Color.Red.A, System.Drawing.Color.Red.R, System.Drawing.Color.Red.G, System.Drawing.Color.Red.B));
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Layer layer = new Layer();
            layer.Name = "Layer1";

            LayerMap_Accessor layerMap = new LayerMap_Accessor(layer);
            layerMap.MapType = LayerMapType.Local;
            layerMap.HeaderRowData = new Collection<string>();
            layerMap.MappedColumnType = new Collection<ColumnType>();
            layerMap.RangeDisplayName = "Sheet_1";
            target.currentLayer = layerMap;
            target.SetDefaultBackground();
            Assert.AreEqual(expected.Color, ((SolidColorBrush)target.ColorBackground).Color);
        }

        /// <summary>
        /// A test for BindDatatoViewModel
        /// </summary>
        [TestMethod()]
        public void BindDatatoViewModelNullTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            target.BindDatatoViewModel();
            Assert.IsTrue(target.IsHelpTextVisible);
            Assert.IsFalse(target.IsTabVisible);
        }

        /// <summary>
        /// A test for BindDatatoViewModel
        /// </summary>
        [TestMethod()]
        public void BindDatatoViewModelTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Layer layer = new Layer();
            layer.Name = "Layer1";

            LayerMap_Accessor layerMap = new LayerMap_Accessor(layer);
            layerMap.MapType = LayerMapType.Local;
            layerMap.HeaderRowData = new Collection<string>();
            layerMap.MappedColumnType = new Collection<ColumnType>();
            layerMap.RangeDisplayName = "Sheet_1";
            target.currentLayer = layerMap;
            target.BindDatatoViewModel();
            Assert.IsTrue(target.IsTabVisible);
            Assert.IsFalse(target.IsHelpTextVisible);
        }

        /// <summary>
        /// A test for RemoveColumns
        /// </summary>
        [TestMethod()]
        public void RemoveColumnsTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Layer layer = new Layer();
            layer.Name = "Layer1";

            LayerMap_Accessor layerMap = new LayerMap_Accessor(layer);
            layerMap.MapType = LayerMapType.Local;
            layerMap.HeaderRowData = new Collection<string>();
            layerMap.MappedColumnType = new Collection<ColumnType>();
            layerMap.RangeDisplayName = "Sheet_1";
            target.currentLayer = layerMap;
            target.selectedGroup = new Group("Sun", GroupType.ReferenceFrame, null);
            Collection<Column> columns = ColumnExtensions.PopulateColumnList();
            Assert.AreEqual(columns.Count, 19);
            target.RemoveColumns(columns);
            Assert.AreEqual(columns.Count, 17);
        }

        /// <summary>
        /// A test for PopulateColumns
        /// </summary>
        [TestMethod()]
        public void PopulateColumnsTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Layer layer = new Layer();
            layer.Name = "Layer1";

            LayerMap_Accessor layerMap = new LayerMap_Accessor(layer);
            layerMap.MapType = LayerMapType.Local;

            layerMap.HeaderRowData = new Collection<string>();
            layerMap.HeaderRowData.Add("Lat");

            layerMap.MappedColumnType = new Collection<ColumnType>();
            layerMap.MappedColumnType.Add(ColumnType.Alt);

            layerMap.RangeDisplayName = "Sheet_1";
            target.currentLayer = layerMap;
            target.selectedGroup = new Group("Sun", GroupType.ReferenceFrame, null);

            target.ColumnsView = new ObservableCollection<ColumnViewModel>();
            target.ColumnsView.Add(new ColumnViewModel() { ExcelHeaderColumn = "Alt", SelectedWWTColumn = new Column(ColumnType.Alt, "Alt", new Collection<string>() { "RA" }) });

            Collection<Column> columns = ColumnExtensions.PopulateColumnList();
            target.PopulateColumns(layerMap, columns);
            Assert.AreEqual(target.ColumnsView.Count, layerMap.HeaderRowData.Count);
            int index = 0;
            foreach (ColumnViewModel column in target.ColumnsView)
            {
                Assert.AreEqual(layerMap.HeaderRowData[index], column.ExcelHeaderColumn);
                Assert.AreEqual(layerMap.MappedColumnType[index], column.SelectedWWTColumn.ColType);
                index++;
            }
        }

        /// <summary>
        /// A test for SetColumnMapping
        /// </summary>
        [TestMethod()]
        public void SetColumnMappingTest()
        {
            LayerDetailsViewModel_Accessor target = new LayerDetailsViewModel_Accessor();
            Layer layer = new Layer();
            layer.Name = "Layer1";

            ObservableCollection<ColumnViewModel> columnsView = new ObservableCollection<ColumnViewModel>();
            columnsView.Add(new ColumnViewModel() { ExcelHeaderColumn = "Alt", SelectedWWTColumn = new Column(ColumnType.Alt, "Alt", new Collection<string>() { "RA" }) });
            columnsView.Add(new ColumnViewModel() { ExcelHeaderColumn = "Dec", SelectedWWTColumn = new Column(ColumnType.None, "Dec", new Collection<string>() { "Dec" }) });

            target.ColumnsView = columnsView;

            foreach (ColumnViewModel columnValue in target.ColumnsView)
            {
                columnValue.WWTColumns = new ObservableCollection<Column>();
                ColumnExtensions.PopulateColumnList().ToList().ForEach(col => columnValue.WWTColumns.Add(col));
            }

            KeyValuePair<int, string> selectedSize = new KeyValuePair<int, string>(1, "Dec");
            target.SetColumnMapping(selectedSize);

            foreach (ColumnViewModel columnView in target.ColumnsView)
            {
                if (columnView.ExcelHeaderColumn == "Dec")
                {
                    Assert.AreEqual(columnView.SelectedWWTColumn.ColType, ColumnType.Mag);
                }
            }
        }
    }
}