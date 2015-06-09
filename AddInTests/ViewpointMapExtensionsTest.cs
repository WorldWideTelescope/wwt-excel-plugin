//-----------------------------------------------------------------------
// <copyright file="ViewpointMapExtensionsTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using Microsoft.Research.Wwt.Excel.Addin;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using InteropExcel = Microsoft.Office.Interop.Excel;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{   
    /// <summary>
    /// This is a test class for ViewpointMapExtensionsTest and is intended
    /// to contain all ViewpointMapExtensionsTest Unit Tests
    /// </summary>
    [TestClass()]
    public class ViewpointMapExtensionsTest
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
        /// A test for Deserialize
        /// </summary>
        [TestMethod()]
        public void DeserializeTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {   
                InteropExcel.Workbook workbook = excelApp.ActiveWorkbook;
                ViewpointMap expected = new ViewpointMap(workbook);
                expected.SerializablePerspective = new System.Collections.ObjectModel.ObservableCollection<Perspective>();
                expected.SerializablePerspective.Add(new Perspective("Earth", "Earth", false, "2/25/2011 7:24:01 AM", "1", "9932 km", "SD8834DFA", "30.0", "30.0", "2.0", ".3", "1.5"));

                ViewpointMap viewpointMap = new ViewpointMap(workbook);
                string xmlContent = "<?xml version=\"1.0\" encoding=\"utf-16\"?><ViewpointMap xmlns:d1p1=\"http://schemas.datacontract.org/2004/07/Microsoft.Research.Wwt.Excel.Addin\" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"Microsoft.Research.Wwt.Excel.Addin.ViewpointMap\"><d1p1:SerializablePerspective xmlns:d2p1=\"http://schemas.datacontract.org/2004/07/Microsoft.Research.Wwt.Excel.Common\"><d2p1:Perspective><d2p1:Declination i:nil=\"true\" /><d2p1:HasRADec>false</d2p1:HasRADec><d2p1:Latitude>2/25/2011 7:24:01 AM</d2p1:Latitude><d2p1:Longitude>1</d2p1:Longitude><d2p1:LookAngle>30.0</d2p1:LookAngle><d2p1:LookAt>Earth</d2p1:LookAt><d2p1:Name i:nil=\"true\" /><d2p1:ObservingTime>30.0</d2p1:ObservingTime><d2p1:ReferenceFrame>Earth</d2p1:ReferenceFrame><d2p1:RightAscention i:nil=\"true\" /><d2p1:Rotation>SD8834DFA</d2p1:Rotation><d2p1:TimeRate>2.0</d2p1:TimeRate><d2p1:ViewToken>1.5</d2p1:ViewToken><d2p1:Zoom>9932 km</d2p1:Zoom><d2p1:ZoomText>.3</d2p1:ZoomText></d2p1:Perspective></d1p1:SerializablePerspective></ViewpointMap>";
                
                ViewpointMap actual;
                actual = ViewpointMapExtensions.Deserialize(viewpointMap, xmlContent);
                Assert.AreEqual(expected.SerializablePerspective[0].Name, actual.SerializablePerspective[0].Name);
            }
            finally
            {
                excelApp.Close();
            }
        }

        /// <summary>
        /// A test for Serialize
        /// </summary>
        [TestMethod()]
        public void SerializeTest()
        {
            InteropExcel.Application excelApp = new InteropExcel.Application();

            try
            {   
                InteropExcel.Workbook workbook = excelApp.ActiveWorkbook;
                ViewpointMap viewpointMap = new ViewpointMap(workbook);
                viewpointMap.SerializablePerspective = new System.Collections.ObjectModel.ObservableCollection<Perspective>();
                viewpointMap.SerializablePerspective.Add(new Perspective("Earth", "Earth", false, "2/25/2011 7:24:01 AM", "1", "9932 km", "SD8834DFA", "30.0", "30.0", "2.0", ".3", "1.5"));
                string expected = "<?xml version=\"1.0\" encoding=\"utf-16\"?><ViewpointMap xmlns:d1p1=\"http://schemas.datacontract.org/2004/07/Microsoft.Research.Wwt.Excel.Addin\" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"Microsoft.Research.Wwt.Excel.Addin.ViewpointMap\"><d1p1:SerializablePerspective xmlns:d2p1=\"http://schemas.datacontract.org/2004/07/Microsoft.Research.Wwt.Excel.Common\"><d2p1:Perspective><d2p1:Declination i:nil=\"true\" /><d2p1:HasRADec>false</d2p1:HasRADec><d2p1:Latitude>2/25/2011 7:24:01 AM</d2p1:Latitude><d2p1:Longitude>1</d2p1:Longitude><d2p1:LookAngle>30.0</d2p1:LookAngle><d2p1:LookAt>Earth</d2p1:LookAt><d2p1:Name i:nil=\"true\" /><d2p1:ObservingTime>30.0</d2p1:ObservingTime><d2p1:ReferenceFrame>Earth</d2p1:ReferenceFrame><d2p1:RightAscention i:nil=\"true\" /><d2p1:Rotation>SD8834DFA</d2p1:Rotation><d2p1:TimeRate>2.0</d2p1:TimeRate><d2p1:ViewToken>1.5</d2p1:ViewToken><d2p1:Zoom>9932 km</d2p1:Zoom><d2p1:ZoomText>.3</d2p1:ZoomText></d2p1:Perspective></d1p1:SerializablePerspective></ViewpointMap>";
                string actual;
                actual = ViewpointMapExtensions.Serialize(viewpointMap);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                excelApp.Close();
            }
        }
    }
}
