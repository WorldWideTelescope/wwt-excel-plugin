//-----------------------------------------------------------------------
// <copyright file="WWTManagerTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// This is a test class for WWTManagerTest and is intended to contain all WWTManagerTest Unit Tests.
    /// </summary>
    [TestClass()]
    public class WWTManagerTest
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
        /// A test for SetMarkerProperties
        /// </summary>
        [TestMethod()]
        public void SetMarkerPropertiesTest()
        {
            // Create an empty layer object.
            Layer layer = new Layer();

            XElement root = XElement.Parse("<LayerApi><Status>Success</Status><Layer Class=\"SpreadSheetLayer\" BeginRange=\"12/31/9999 11:59:59 PM\" EndRange=\"1/1/0001 12:00:00 AM\" Decay=\"16\" CoordinatesType=\"Spherical\" LatColumn=\"-1\" LngColumn=\"-1\" GeometryColumn=\"-1\" XAxisColumn=\"-1\" YAxisColumn=\"6\" ZAxisColumn=\"-1\" XAxisReverse=\"False\" YAxisReverse=\"False\" ZAxisReverse=\"False\" AltType=\"Depth\" MarkerMix=\"Same_For_All\" MarkerColumn=\"-1\" ColorMapColumn=\"-1\" PlotType=\"Gaussian\" MarkerIndex=\"0\" ShowFarSide=\"False\" MarkerScale=\"World\" AltUnit=\"Meters\" CartesianScale=\"Meters\" CartesianCustomScale=\"1\" AltColumn=\"-1\" StartDateColumn=\"-1\" EndDateColumn=\"-1\" SizeColumn=\"-1\" NameColumn=\"-1\" HyperlinkFormat=\"\" HyperlinkColumn=\"-1\" ScaleFactor=\"1\" PointScaleType=\"Power\" Opacity=\"1\" StartTime=\"1/1/0001 12:00:00 AM\" EndTime=\"12/31/9999 11:59:59 PM\" FadeSpan=\"00:00:00\" FadeType=\"None\" Name=\"Sheet1_1\" ColorValue=\"ARGBColor:255:255:0:0\" Enabled=\"True\" Astronomical=\"False\" /></LayerApi>", LoadOptions.PreserveWhitespace);

            // Get All Attributes list  of Layers.
            var listOfAttributes = root.Element(Constants.LayerElementNodeName).Attributes();

            WWTManager_Accessor.SetMarkerProperties(layer, listOfAttributes);

            Assert.AreEqual(layer.TimeDecay, 16);
            Assert.AreEqual(layer.ScaleFactor, 1);
            Assert.AreEqual(layer.Opacity, 1);
            Assert.AreEqual(layer.StartTime, Convert.ToDateTime("1/1/0001 12:00:00 AM", CultureInfo.CurrentCulture));
            Assert.AreEqual(layer.EndTime, Convert.ToDateTime("12/31/9999 11:59:59 PM", CultureInfo.CurrentCulture));
            Assert.AreEqual(layer.FadeSpan, TimeSpan.Parse("00:00:00", CultureInfo.CurrentCulture));
            Assert.AreEqual(layer.Color, "ARGBColor:255:255:0:0");
            Assert.AreEqual(layer.AltType, AltType.Depth);
            Assert.AreEqual(layer.MarkerScale, ScaleRelativeType.World);
            Assert.AreEqual(layer.AltUnit, AltUnit.Meters);
            Assert.AreEqual(layer.PointScaleType, ScaleType.Power);
            Assert.AreEqual(layer.FadeType, FadeType.None);
        }

        /// <summary>
        /// A test for GetLayerProperties
        /// </summary>
        [TestMethod()]
        public void GetLayerPropertiesLatLongTest()
        {
            string expected = "<LayerApi><Layer Name=\"California\" CoordinatesType=\"Spherical\" XAxisColumn=\"-1\" YAxisColumn=\"-1\" ZAxisColumn=\"-1\" XAxisReverse=\"false\" YAxisReverse=\"false\" ZAxisReverse=\"false\" LatColumn=\"2\" LngColumn=\"1\" GeometryColumn=\"-1\" ColorMapColumn=\"6\" AltColumn=\"3\" StartDateColumn=\"4\" EndDateColumn=\"-1\" SizeColumn=\"5\" NameColumn=\"0\" Decay=\"0\" ScaleFactor=\"1\" Opacity=\"0\" StartTime=\"1/1/0001 12:00:00 AM\" EndTime=\"12/31/9999 11:59:59 PM\" FadeSpan=\"00:00:00\" ColorValue=\"ARGBColor:255:255:255:255\" AltType=\"Altitude\" MarkerScale=\"World\" AltUnit=\"Meters\" RaUnits=\"Hours\" PointScaleType=\"Power\" FadeType=\"None\" PlotType=\"Gaussian\" MarkerIndex=\"0\" ShowFarSide=\"true\" /></LayerApi>";

            Layer layer = new Layer();
            
            layer.Name = "California";
            layer.LatColumn = 2;
            layer.LngColumn = 1;
            layer.GeometryColumn = -1;
            layer.ColorMapColumn = 6;
            layer.AltColumn = 3;
            layer.StartDateColumn = 4;
            layer.EndDateColumn = -1;
            layer.SizeColumn = 5;
            layer.NameColumn = 0;
            layer.TimeDecay = 0;
            layer.ScaleFactor = 1;
            layer.Opacity = 0;
            layer.StartTime = Convert.ToDateTime("1/1/0001 12:00:00 AM", CultureInfo.InvariantCulture);
            layer.EndTime = Convert.ToDateTime("12/31/9999 11:59:59 PM", CultureInfo.InvariantCulture);
            layer.FadeSpan = new TimeSpan();
            layer.Color = "ARGBColor:255:255:255:255";
            layer.AltType = AltType.Altitude;
            layer.MarkerScale = ScaleRelativeType.World;
            layer.AltUnit = AltUnit.Meters;
            layer.PointScaleType = ScaleType.Power;
            layer.FadeType = FadeType.None;

            // Dummy sun group
            Group group = new Group("Sun", GroupType.ReferenceFrame, null);
            layer.Group = group;

            string actual = WWTManager_Accessor.GetLayerProperties(layer, false);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for GetLayerProperties
        /// </summary>
        [TestMethod()]
        public void GetLayerPropertiesRaDecTest()
        {
            string expected = "<LayerApi><Layer Name=\"California\" CoordinatesType=\"Spherical\" XAxisColumn=\"-1\" YAxisColumn=\"-1\" ZAxisColumn=\"-1\" XAxisReverse=\"false\" YAxisReverse=\"false\" ZAxisReverse=\"false\" LatColumn=\"2\" LngColumn=\"1\" GeometryColumn=\"-1\" ColorMapColumn=\"6\" AltColumn=\"3\" StartDateColumn=\"4\" EndDateColumn=\"-1\" SizeColumn=\"5\" NameColumn=\"0\" Decay=\"0\" ScaleFactor=\"1\" Opacity=\"0\" StartTime=\"1/1/0001 12:00:00 AM\" EndTime=\"12/31/9999 11:59:59 PM\" FadeSpan=\"00:00:00\" ColorValue=\"ARGBColor:255:255:255:255\" AltType=\"Altitude\" MarkerScale=\"World\" AltUnit=\"Meters\" RaUnits=\"Hours\" PointScaleType=\"Power\" FadeType=\"None\" PlotType=\"Gaussian\" MarkerIndex=\"0\" ShowFarSide=\"true\" /></LayerApi>";

            Layer layer = new Layer();

            layer.Name = "California";
            layer.RAColumn = 1;
            layer.DecColumn = 2;
            layer.GeometryColumn = -1;
            layer.ColorMapColumn = 6;
            layer.AltColumn = 3;
            layer.StartDateColumn = 4;
            layer.EndDateColumn = -1;
            layer.SizeColumn = 5;
            layer.NameColumn = 0;
            layer.TimeDecay = 0;
            layer.ScaleFactor = 1;
            layer.Opacity = 0;
            layer.StartTime = Convert.ToDateTime("1/1/0001 12:00:00 AM", CultureInfo.InvariantCulture);
            layer.EndTime = Convert.ToDateTime("12/31/9999 11:59:59 PM", CultureInfo.InvariantCulture);
            layer.FadeSpan = new TimeSpan();
            layer.Color = "ARGBColor:255:255:255:255";
            layer.AltType = AltType.Altitude;
            layer.MarkerScale = ScaleRelativeType.World;
            layer.AltUnit = AltUnit.Meters;
            layer.PointScaleType = ScaleType.Power;
            layer.FadeType = FadeType.None;

            // Dummy sun group
            Group group = new Group("Sky", GroupType.ReferenceFrame, null);
            layer.Group = group;

            string actual = WWTManager_Accessor.GetLayerProperties(layer, false);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for SplitString
        /// </summary>
        [TestMethod()]
        public void SplitStringTest()
        {
            string data = "Letter\tLong\tLAT\tAltitude\tdatetime\tMagnitude\tColor";
            List<string> expected = new List<string> { "Letter", "Long", "LAT", "Altitude", "datetime", "Magnitude", "Color" };
            List<string> actual = WWTManager_Accessor.SplitString(data);
            Assert.AreEqual(expected.Count, actual.Count);
            Assert.AreEqual(expected[0], actual[0]);
            Assert.AreEqual(expected[1], actual[1]);
            Assert.AreEqual(expected[2], actual[2]);
            Assert.AreEqual(expected[3], actual[3]);
            Assert.AreEqual(expected[4], actual[4]);
            Assert.AreEqual(expected[5], actual[5]);
            Assert.AreEqual(expected[6], actual[6]);
        }

        /// <summary>
        /// A test for IsValidMachine
        /// </summary>
        [TestMethod()]
        public void IsValidMachineTest()
        {
            // User WWTMockRequest in WWTManager so that all calls to WWT API will succeed.
            Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());

            bool actual = WWTManager_Accessor.IsValidMachine(TargetMachine.DefaultIP.ToString(), true);
            Assert.AreEqual(actual, true);
        }

        /// <summary>
        /// A test for SetLayerProperties with Lat/Long columns
        /// </summary>
        [TestMethod()]
        public void SetLayerPropertiesWithLatLongTest()
        {
            Layer layer = new Layer();
            layer.Group = new Group("Sun", GroupType.ReferenceFrame, null);

            string expected = "<LayerApi><Layer Name=\"California\" CoordinatesType=\"Spherical\" XAxisColumn=\"-1\" YAxisColumn=\"-1\" ZAxisColumn=\"-1\" XAxisReverse=\"false\" YAxisReverse=\"false\" ZAxisReverse=\"false\" LatColumn=\"2\" LngColumn=\"1\" GeometryColumn=\"-1\" ColorMapColumn=\"6\" AltColumn=\"3\" StartDateColumn=\"4\" EndDateColumn=\"-1\" SizeColumn=\"5\" NameColumn=\"0\" Decay=\"16\" ScaleFactor=\"8\" Opacity=\"1\" StartTime=\"1/1/0001 12:00:00 AM\" EndTime=\"12/31/9999 11:59:59 PM\" FadeSpan=\"00:00:00\" ColorValue=\"ARGBColor:255:255:0:0\" AltType=\"Depth\" MarkerScale=\"World\" AltUnit=\"Meters\" RaUnits=\"Hours\" PointScaleType=\"Power\" FadeType=\"None\" PlotType=\"Gaussian\" MarkerIndex=\"0\" ShowFarSide=\"true\" /></LayerApi>";

            XElement element = XElement.Parse(expected);
            var listOfAttributes = element.Element(Constants.LayerElementNodeName).Attributes();
            WWTManager_Accessor.SetLayerProperties(layer, listOfAttributes);

            string actual = WWTManager_Accessor.GetLayerProperties(layer, false);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for SetLayerProperties with RA/Dec columns
        /// </summary>
        [TestMethod()]
        public void SetLayerPropertiesWithRaDecTest()
        {
            Layer layer = new Layer();
            layer.Group = new Group("Sky", GroupType.ReferenceFrame, null);

            string expected = "<LayerApi><Layer Name=\"California\" CoordinatesType=\"Spherical\" XAxisColumn=\"-1\" YAxisColumn=\"-1\" ZAxisColumn=\"-1\" XAxisReverse=\"false\" YAxisReverse=\"false\" ZAxisReverse=\"false\" LatColumn=\"2\" LngColumn=\"1\" GeometryColumn=\"-1\" ColorMapColumn=\"6\" AltColumn=\"3\" StartDateColumn=\"4\" EndDateColumn=\"-1\" SizeColumn=\"5\" NameColumn=\"0\" Decay=\"16\" ScaleFactor=\"8\" Opacity=\"1\" StartTime=\"1/1/0001 12:00:00 AM\" EndTime=\"12/31/9999 11:59:59 PM\" FadeSpan=\"00:00:00\" ColorValue=\"ARGBColor:255:255:0:0\" AltType=\"Depth\" MarkerScale=\"World\" AltUnit=\"Meters\" RaUnits=\"Hours\" PointScaleType=\"Power\" FadeType=\"None\" PlotType=\"Gaussian\" MarkerIndex=\"0\" ShowFarSide=\"true\" /></LayerApi>";

            XElement element = XElement.Parse(expected);
            var listOfAttributes = element.Element(Constants.LayerElementNodeName).Attributes();
            WWTManager_Accessor.SetLayerProperties(layer, listOfAttributes);

            string actual = WWTManager_Accessor.GetLayerProperties(layer, false);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsValidLayer
        /// </summary>
        [TestMethod()]
        public void IsValidLayerTest()
        {
            // Use WWTMockRequest in WWTManager so that all calls to WWT API will succeed.
            Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
            Globals_Accessor.TargetMachine = new TargetMachine(Constants.Localhost);
            string layerId = "c71ebb83-a2b2-437b-8cef-1524b4c8aa7e";
            bool expected = true;
            bool actual = WWTManager_Accessor.IsValidLayer(layerId);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsValidLayer
        /// </summary>
        [TestMethod()]
        public void IsValidLayerNegativeTest()
        {
            // Use WWTMockRequest in WWTManager so that all calls to WWT API will succeed.
            Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
            Globals_Accessor.TargetMachine = new TargetMachine(Constants.Localhost);
            string layerId = "InvalidID";
            bool expected = false;
            bool actual = WWTManager_Accessor.IsValidLayer(layerId);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsValidGroup
        /// </summary>
        [TestMethod()]
        public void IsValidGroupTest()
        {
            // Use WWTMockRequest in WWTManager so that all calls to WWT API will succeed.
            Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
            Globals_Accessor.TargetMachine = new TargetMachine(Constants.Localhost);
            Group group = new Group("Moon", GroupType.ReferenceFrame, new Group("Earth", GroupType.ReferenceFrame, null));
            ICollection<Group> wwtGroups = WWTManager_Accessor.GetAllWWTGroups(true);
            bool expected = true;
            bool actual;
            actual = WWTManager.IsValidGroup(group, wwtGroups);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsValidGroup
        /// </summary>
        [TestMethod()]
        public void IsValidGroupNegativeTest()
        {
            // Use WWTMockRequest in WWTManager so that all calls to WWT API will succeed.
            Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
            Globals_Accessor.TargetMachine = new TargetMachine(Constants.Localhost);
            Group group = new Group("Insat", GroupType.ReferenceFrame, new Group("Earth", GroupType.ReferenceFrame, null));
            ICollection<Group> wwtGroups = WWTManager_Accessor.GetAllWWTGroups(true);
            bool expected = false;
            bool actual;
            actual = WWTManager.IsValidGroup(group, wwtGroups);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for UploadDataInWWT
        /// </summary>
        [TestMethod()]
        public void UploadDataInWWTTest()
        {
            // Use WWTMockRequest in WWTManager so that all calls to WWT API will succeed.
            Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
            Globals_Accessor.TargetMachine = new TargetMachine(Constants.Localhost);
            
            string[] data = new string[] { "TestData\r\n" };

            string layerId = "3f0cfda2-7319-4190-8f5e-99778a04ca3d";
            bool isConsumeException = true;
            bool expected = true;
            bool actual;
            actual = WWTManager.UploadDataInWWT(layerId, data, isConsumeException);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for CreateLayer
        /// </summary>
        [TestMethod()]
        public void CreateLayerTest()
        {
            // Use WWTMockRequest in WWTManager so that all calls to WWT API will succeed.
            Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());
            Globals_Accessor.TargetMachine = new TargetMachine(Constants.Localhost);

            string headerData = "LAT\tLON\tDepth\tTime\tContributor\tRegion\tColor\tMag";
            string expected = "3f0cfda2-7319-4190-8f5e-99778a04ca3d";
            string actual = WWTManager_Accessor.CreateLayer("TestLayer", "Earth", headerData);

            Assert.AreEqual(expected, actual);
        }
    }
}