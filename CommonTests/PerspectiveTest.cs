//-----------------------------------------------------------------------
// <copyright file="PerspectiveTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{   
    /// <summary>
    /// This is a test class for PerspectiveTest and is intended
    /// to contain all PerspectiveTest Unit Tests
    /// </summary>
    [TestClass()]
    public class PerspectiveTest
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
        /// A test for Perspective Constructor
        /// </summary>
        [TestMethod()]
        public void PerspectiveConstructorTest()
        {
            string lookAt = "Earth";
            string referenceFrame = "Earth";
            bool hasRADec = false;
            string observingTime = "2/25/2011 7:24:01 AM";
            string timeRate = "1";
            string zoomText = "9932 km";
            string viewToken = "SD8834DFA";
            string latitude = "30.0";
            string longitude = "30.0";
            string zoom = "2.0";
            string rotation = ".3";
            string lookAngle = "1.5";
            Perspective target = new Perspective(lookAt, referenceFrame, hasRADec, latitude, longitude, zoom, rotation, lookAngle, observingTime, timeRate, zoomText, viewToken);
            Assert.IsNotNull(target);
        }

        /// <summary>
        /// A test for Latitude
        /// </summary>
        [TestMethod()]
        public void LatitudeTest()
        {
            string lookAt = "Earth";
            string referenceFrame = "Earth";
            bool hasRADec = false;
            string observingTime = "2/25/2011 7:24:01 AM";
            string timeRate = "1";
            string zoomText = "9932 km";
            string viewToken = "SD8834DFA";
            string latitude = "30.0";
            string longitude = "30.0";
            string zoom = "2.0";
            string rotation = ".3";
            string lookAngle = "1.5";
            Perspective target = new Perspective(lookAt, referenceFrame, hasRADec, latitude, longitude, zoom, rotation, lookAngle, observingTime, timeRate, zoomText, viewToken);
            string expected = latitude;
            string actual = target.Latitude;
            Assert.AreEqual(expected, actual);
         }

        /// <summary>
        /// A test for Longitude
        /// </summary>
        [TestMethod()]
        public void LongitudeTest()
        {
            string lookAt = "Earth";
            string referenceFrame = "Earth";
            bool hasRADec = false;
            string observingTime = "2/25/2011 7:24:01 AM";
            string timeRate = "1";
            string zoomText = "9932 km";
            string viewToken = "SD8834DFA";
            string latitude = "30.0";
            string longitude = "30.0";
            string zoom = "2.0";
            string rotation = ".3";
            string lookAngle = "1.5";
            Perspective target = new Perspective(lookAt, referenceFrame, hasRADec, latitude, longitude, zoom, rotation, lookAngle, observingTime, timeRate, zoomText, viewToken);
            string expected = longitude;
            string actual = target.Longitude;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for LookAngle
        /// </summary>
        [TestMethod()]
        public void LookAngleTest()
        {
            string lookAt = "Earth";
            string referenceFrame = "Earth";
            bool hasRADec = false;
            string observingTime = "2/25/2011 7:24:01 AM";
            string timeRate = "1";
            string zoomText = "9932 km";
            string viewToken = "SD8834DFA";
            string latitude = "30.0";
            string longitude = "30.0";
            string zoom = "2.0";
            string rotation = ".3";
            string lookAngle = "1.5";
            Perspective target = new Perspective(lookAt, referenceFrame, hasRADec, latitude, longitude, zoom, rotation, lookAngle, observingTime, timeRate, zoomText, viewToken);
            string expected = lookAngle;
            string actual = target.LookAngle;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for Rotation
        /// </summary>
        [TestMethod()]
        public void RotationTest()
        {
            string lookAt = "Earth";
            string referenceFrame = "Earth";
            bool hasRADec = false;
            string observingTime = "2/25/2011 7:24:01 AM";
            string timeRate = "1";
            string zoomText = "9932 km";
            string viewToken = "SD8834DFA";
            string latitude = "30.0";
            string longitude = "30.0";
            string zoom = "2.0";
            string rotation = ".3";
            string lookAngle = "1.5";
            Perspective target = new Perspective(lookAt, referenceFrame, hasRADec, latitude, longitude, zoom, rotation, lookAngle, observingTime, timeRate, zoomText, viewToken);
            string expected = rotation;
            string actual = target.Rotation;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for Zoom
        /// </summary>
        [TestMethod()]
        public void ZoomTest()
        {
            string lookAt = "Earth";
            string referenceFrame = "Earth";
            bool hasRADec = false;
            string observingTime = "2/25/2011 7:24:01 AM";
            string timeRate = "1";
            string zoomText = "9932 km";
            string viewToken = "SD8834DFA";
            string latitude = "30.0";
            string longitude = "30.0";
            string zoom = "2.0";
            string rotation = ".3";
            string lookAngle = "1.5";
            Perspective target = new Perspective(lookAt, referenceFrame, hasRADec, latitude, longitude, zoom, rotation, lookAngle, observingTime, timeRate, zoomText, viewToken);
            string expected = zoom;
            string actual = target.Zoom;
            Assert.AreEqual(expected, actual);
        }
    }
}
