//-----------------------------------------------------------------------
// <copyright file="LayerTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// This is a test class for LayerTest and is intended
    /// to contain all LayerTest Unit Tests
    /// </summary>
    [TestClass()]
    public class LayerTest
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
        /// A test for Layer Constructor
        /// </summary>
        [TestMethod()]
        public void LayerConstructorTest()
        {
            Layer target = new Layer();
            Assert.IsNotNull(target);
        }

        /// <summary>
        /// A test for InitilizeColumnDefaults
        /// </summary>
        [TestMethod()]
        public void InitializeColumnDefaultsTest()
        {
            Layer target = new Layer(); // TODO: Initialize to an appropriate value
            target.InitializeColumnDefaults();
            Assert.IsNotNull(target);
        }

        /// <summary>
        /// A test for InitilizeDefaults
        /// </summary>
        [TestMethod()]
        public void InitializeDefaultsTest()
        {
            Layer target = new Layer();
            target.InitializeDefaults();
            Assert.IsNotNull(target);
        }

        /// <summary>
        /// A test for AltColumn
        /// </summary>
        [TestMethod()]
        public void AltColumnTest()
        {
            Layer target = new Layer();
            int expected = 3;
            target.AltColumn = expected;
            int actual = target.AltColumn;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for AltType
        /// </summary>
        [TestMethod()]
        public void AltTypeTest()
        {
            Layer target = new Layer();
            AltType expected = AltType.Depth;
            target.AltType = expected;
            AltType actual = target.AltType;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for AltUnit
        /// </summary>
        [TestMethod()]
        public void AltUnitTest()
        {
            Layer target = new Layer();
            AltUnit expected = AltUnit.Meters;
            target.AltUnit = expected;
            AltUnit actual = target.AltUnit;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for Color
        /// </summary>
        [TestMethod()]
        public void ColorTest()
        {
            Layer target = new Layer();
            string expected = "Red";
            target.Color = expected;
            string actual = target.Color;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for ColorMapColumn
        /// </summary>
        [TestMethod()]
        public void ColorMapColumnTest()
        {
            Layer target = new Layer();
            int expected = 4;
            target.ColorMapColumn = expected;
            int actual = target.ColorMapColumn;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for EndDateColumn
        /// </summary>
        [TestMethod()]
        public void EndDateColumnTest()
        {
            Layer target = new Layer();
            int expected = 24;
            int actual;
            target.EndDateColumn = expected;
            actual = target.EndDateColumn;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for EndTime
        /// </summary>
        [TestMethod()]
        public void EndTimeTest()
        {
            Layer target = new Layer();
            DateTime expected = DateTime.Now;
            target.EndTime = expected;
            DateTime actual = target.EndTime;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for FadeSpan
        /// </summary>
        [TestMethod()]
        public void FadeSpanTest()
        {
            Layer target = new Layer();
            TimeSpan expected = new TimeSpan(0, 0, 10);
            target.FadeSpan = expected;
            TimeSpan actual = target.FadeSpan;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for FadeType
        /// </summary>
        [TestMethod()]
        public void FadeTypeTest()
        {
            Layer target = new Layer();
            FadeType expected = FadeType.In;
            target.FadeType = expected;
            FadeType actual = target.FadeType;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for GeometryColumn
        /// </summary>
        [TestMethod()]
        public void GeometryColumnTest()
        {
            Layer target = new Layer();
            int expected = 92;
            target.GeometryColumn = expected;
            int actual = target.GeometryColumn;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for ID
        /// </summary>
        [TestMethod()]
        public void IDTest()
        {
            Layer target = new Layer();
            string expected = "1234-5678-90ab-cdef";
            target.ID = expected;
            string actual = target.ID;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for LatColumn
        /// </summary>
        [TestMethod()]
        public void LatColumnTest()
        {
            Layer target = new Layer();
            int expected = 6;
            target.LatColumn = expected;
            int actual = target.LatColumn;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for LngColumn
        /// </summary>
        [TestMethod()]
        public void LngColumnTest()
        {
            Layer target = new Layer();
            int expected = 8;
            target.LngColumn = expected;
            int actual = target.LngColumn;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for MarkerScale
        /// </summary>
        [TestMethod()]
        public void MarkerScaleTest()
        {
            Layer target = new Layer();
            ScaleRelativeType expected = ScaleRelativeType.World;
            target.MarkerScale = expected;
            ScaleRelativeType actual = target.MarkerScale;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for Name
        /// </summary>
        [TestMethod()]
        public void NameTest()
        {
            Layer target = new Layer();
            string expected = "New Layer";
            target.Name = expected;
            string actual = target.Name;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for NameColumn
        /// </summary>
        [TestMethod()]
        public void NameColumnTest()
        {
            Layer target = new Layer();
            int expected = 3;
            target.NameColumn = expected;
            int actual = target.NameColumn;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for Opacity
        /// </summary>
        [TestMethod()]
        public void OpacityTest()
        {
            Layer target = new Layer();
            double expected = .5F;
            target.Opacity = expected;
            double actual = target.Opacity;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for PointScaleType
        /// </summary>
        [TestMethod()]
        public void PointScaleTypeTest()
        {
            Layer target = new Layer();
            ScaleType expected = ScaleType.Linear;
            target.PointScaleType = expected;
            ScaleType actual = target.PointScaleType;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for ScaleFactor
        /// </summary>
        [TestMethod()]
        public void ScaleFactorTest()
        {
            Layer target = new Layer();
            double expected = 10F;
            target.ScaleFactor = expected;
            double actual = target.ScaleFactor;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for ScaleFactor Zero
        /// </summary>
        [TestMethod()]
        public void ScaleFactorZeroTest()
        {
            Layer target = new Layer();
            double expected = 0F;
            target.ScaleFactor = expected;
            double actual = target.ScaleFactor;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for Negative ScaleFactor
        /// </summary>
        [TestMethod()]
        public void ScaleFactorNegativeTest()
        {
            Layer target = new Layer();
            double expected = -10F;
            target.ScaleFactor = expected;
            double actual = target.ScaleFactor;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for SizeColumn
        /// </summary>
        [TestMethod()]
        public void SizeColumnTest()
        {
            Layer target = new Layer();
            int expected = 8;
            target.SizeColumn = expected;
            int actual = target.SizeColumn;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for StartDateColumn
        /// </summary>
        [TestMethod()]
        public void StartDateColumnTest()
        {
            Layer target = new Layer();
            int expected = 7;
            target.StartDateColumn = expected;
            int actual = target.StartDateColumn;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for StartTime
        /// </summary>
        [TestMethod()]
        public void StartTimeTest()
        {
            Layer target = new Layer();
            DateTime expected = DateTime.Now;
            target.StartTime = expected;
            DateTime actual = target.StartTime;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for TimeDecay
        /// </summary>
        [TestMethod()]
        public void TimeDecayTest()
        {
            Layer target = new Layer();
            double expected = 3F;
            target.TimeDecay = expected;
            double actual = target.TimeDecay;
            Assert.AreEqual(expected, actual);
        }
    }
}
