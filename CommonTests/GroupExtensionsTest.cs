//-----------------------------------------------------------------------
// <copyright file="GroupExtensionsTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// This is a test class for GroupExtensionsTest and is intended
    /// to contain all GroupExtensionsTest Unit Tests
    /// </summary>
    [TestClass()]
    public class GroupExtensionsTest
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

        #region SerializeTest

        /// <summary>
        /// A test for Serialize
        /// </summary>
        [TestMethod()]
        public void SerializeTest()
        {
            Group parent = new Group("Sun", GroupType.ReferenceFrame, null);
            Group group = new Group("Earth", GroupType.ReferenceFrame, parent);

            string expected = "<?xml version=\"1.0\" encoding=\"utf-16\"?><Group xmlns:d1p1=\"http://schemas.datacontract.org/2004/07/Microsoft.Research.Wwt.Excel.Common\" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"Microsoft.Research.Wwt.Excel.Common\"><d1p1:GroupType>ReferenceFrame</d1p1:GroupType><d1p1:Name>Earth</d1p1:Name><d1p1:Parent><d1p1:GroupType>ReferenceFrame</d1p1:GroupType><d1p1:Name>Sun</d1p1:Name><d1p1:Parent i:nil=\"true\" /><d1p1:Path>/Sun</d1p1:Path></d1p1:Parent><d1p1:Path>/Sun/Earth</d1p1:Path></Group>";
            string actual;
            actual = GroupExtensions.Serialize(group);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for Serialize
        /// </summary>
        [TestMethod()]
        public void SerializeTestNullGroup()
        {
            Group group = null;
            string expected = string.Empty;
            string actual;
            actual = GroupExtensions.Serialize(group);
            Assert.AreEqual(expected, actual);
        }

        #endregion

        #region DeserializeTest

        /// <summary>
        /// A test for Deserialize
        /// </summary>
        [TestMethod()]
        public void DeserializeTest()
        {
            string xmlContent = "<?xml version=\"1.0\" encoding=\"utf-16\"?><Group xmlns:d1p1=\"http://schemas.datacontract.org/2004/07/Microsoft.Research.Wwt.Excel.Common\" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"Microsoft.Research.Wwt.Excel.Common\"><d1p1:GroupType>ReferenceFrame</d1p1:GroupType><d1p1:Name>Earth</d1p1:Name><d1p1:Parent><d1p1:GroupType>ReferenceFrame</d1p1:GroupType><d1p1:Name>Sun</d1p1:Name><d1p1:Parent i:nil=\"true\" /><d1p1:Path>/Sun</d1p1:Path></d1p1:Parent><d1p1:Path>/Sun/Earth</d1p1:Path></Group>";

            Group parent = new Group("Sun", GroupType.ReferenceFrame, null);
            Group expected = new Group("Earth", GroupType.ReferenceFrame, parent);

            Group group = null;
            Group actual;
            actual = GroupExtensions.Deserialize(group, xmlContent);
            Assert.AreEqual(expected.Name, actual.Name);
            Assert.AreEqual(expected.Path, actual.Path);
        }

        /// <summary>
        /// A test for Deserialize
        /// </summary>
        [TestMethod()]
        public void DeserializeTestEmptyStringGroup()
        {
            Group group = null;

            Group actual;
            actual = GroupExtensions.Deserialize(group, string.Empty);
            Assert.IsNull(actual);
        }

        /// <summary>
        /// A test for Deserialize
        /// </summary>
        [TestMethod()]
        public void DeserializeTestNullGroup()
        {
            Group group = null;

            Group actual;
            actual = GroupExtensions.Deserialize(group, null);
            Assert.IsNull(actual);
        }

        #endregion

        #region GetReferenceFrameTest

        /// <summary>
        /// A test for GetReferenceFrame for sun.
        /// </summary>
        [TestMethod()]
        public void GetReferenceFrameTestSun()
        {
            Group group = new Group("Sun", GroupType.ReferenceFrame, null);
            string expected = "Sun";
            string actual;
            actual = GroupExtensions.GetReferenceFrame(group);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for GetReferenceFrame for earth.
        /// </summary>
        [TestMethod()]
        public void GetReferenceFrameTestEarth()
        {
            Group sun = new Group("Sun", GroupType.ReferenceFrame, null);
            Group earth = new Group("Earth", GroupType.ReferenceFrame, sun);
            string expected = "Earth";
            string actual;
            actual = GroupExtensions.GetReferenceFrame(earth);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for GetReferenceFrame for a layer group under earth.
        /// </summary>
        [TestMethod()]
        public void GetReferenceFrameTestLayerGroup()
        {
            Group sun = new Group("Sun", GroupType.ReferenceFrame, null);
            Group earth = new Group("Earth", GroupType.ReferenceFrame, sun);
            Group layerGroup = new Group("LG", GroupType.LayerGroup, earth);
            string expected = "Earth";
            string actual;
            actual = GroupExtensions.GetReferenceFrame(layerGroup);
            Assert.AreEqual(expected, actual);
        }

        #endregion

        #region IsPlanetTest

        /// <summary>
        /// A test for IsPlanet, if the group is earth.
        /// </summary>
        [TestMethod()]
        public void IsPlanetTestEarth()
        {
            Group sun = new Group("Sun", GroupType.ReferenceFrame, null);
            Group earth = new Group("Earth", GroupType.ReferenceFrame, sun);

            bool expected = true;
            bool actual;
            actual = GroupExtensions.IsPlanet(earth);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsPlanet, if the group is sky.
        /// </summary>
        [TestMethod()]
        public void IsPlanetTestSky()
        {
            Group sky = new Group("sky", GroupType.ReferenceFrame, null);

            bool expected = false;
            bool actual;
            actual = GroupExtensions.IsPlanet(sky);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsPlanet, if the value of group is null.
        /// </summary>
        [TestMethod()]
        public void IsPlanetTestNull()
        {
            bool expected = false;
            bool actual;
            actual = GroupExtensions.IsPlanet(null);
            Assert.AreEqual(expected, actual);
        }

        #endregion

        #region GetDefaultEarthGroupTest

        /// <summary>
        /// A test for GetDefaultGroup when there are no groups.
        /// </summary>
        [TestMethod()]
        public void GetDefaultEarthGroupTestNoGroups()
        {
            ICollection<Group> groups = new List<Group>();
            Group actual;
            actual = GroupExtensions.GetDefaultEarthGroup(groups);
            Assert.AreEqual(actual.Name, "Earth");
            Assert.AreEqual(actual.Path, "/Sun/Earth");
            Assert.AreEqual(actual.IsDeleted, true);
        }

        /// <summary>
        /// A test for GetDefaultGroup when there are groups.
        /// </summary>
        [TestMethod()]
        public void GetDefaultEarthGroupTestWithGroups()
        {
            ICollection<Group> groups = GetGroups();
            Group actual;
            actual = GroupExtensions.GetDefaultEarthGroup(groups);
            Assert.AreEqual(actual.Name, "Earth");
            Assert.AreEqual(actual.Path, "/Sun/Earth");
            Assert.AreEqual(actual.IsDeleted, false);
        }

        #endregion

        #region GetDefaultSkyGroupTest

        /// <summary>
        /// A test for GetDefaultGroup when there are no groups.
        /// </summary>
        [TestMethod()]
        public void GetDefaultSkyGroupTestNoGroups()
        {
            ICollection<Group> groups = new List<Group>();
            Group actual;
            actual = GroupExtensions.GetDefaultSkyGroup(groups);
            Assert.AreEqual(actual.Name, "Sky");
            Assert.AreEqual(actual.Path, "/Sky");
            Assert.AreEqual(actual.IsDeleted, true);
        }

        /// <summary>
        /// A test for GetDefaultGroup when there are groups.
        /// </summary>
        [TestMethod()]
        public void GetDefaultSkyGroupTestWithGroups()
        {
            ICollection<Group> groups = GetGroups();
            Group actual;
            actual = GroupExtensions.GetDefaultSkyGroup(groups);
            Assert.AreEqual(actual.Name, "Sky");
            Assert.AreEqual(actual.Path, "/Sky");
            Assert.AreEqual(actual.IsDeleted, false);
        }

        #endregion

        #region SearchGroupTest

        /// <summary>
        /// A test for SearchGroup
        /// </summary>
        [TestMethod()]
        public void SearchGroupTestNull()
        {
            string groupName = string.Empty;
            ICollection<Group> wwtGroups = new List<Group>();
            Group actual;
            actual = GroupExtensions.SearchGroup(wwtGroups, groupName);
            Assert.IsNull(actual);
        }

        /// <summary>
        /// A test for SearchGroup with wrong name.
        /// </summary>
        [TestMethod()]
        public void SearchGroupTestWrongName()
        {
            string groupName = "WrongName";
            ICollection<Group> wwtGroups = GetGroups();
            Group expected = null;
            Group actual;
            actual = GroupExtensions.SearchGroup(wwtGroups, groupName);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for SearchGroup with wrong name.
        /// </summary>
        [TestMethod()]
        public void SearchGroupTestValid()
        {
            string groupName = "Moon";
            ICollection<Group> wwtGroups = GetGroups();
            Group actual;
            actual = GroupExtensions.SearchGroup(wwtGroups, groupName);
            Assert.AreEqual(actual.GroupType, GroupType.ReferenceFrame);
            Assert.AreEqual(actual.Name, "Moon");
            Assert.AreEqual(actual.Path, "/Sun/Earth/Moon");
        }

        #endregion

        private static ICollection<Group> GetGroups()
        {
            ICollection<Group> groups = new List<Group>();

            Group sun = new Group("Sun", GroupType.ReferenceFrame, null);
            Group earth = new Group("Earth", GroupType.ReferenceFrame, sun);
            Group moon = new Group("Moon", GroupType.ReferenceFrame, earth);
            Group mars = new Group("Mars", GroupType.ReferenceFrame, sun);
            Group jupiter = new Group("Jupiter", GroupType.ReferenceFrame, sun);
            Group saturn = new Group("Saturn", GroupType.ReferenceFrame, sun);
            Group uranus = new Group("Uranus", GroupType.ReferenceFrame, sun);

            earth.Children.Add(moon);

            sun.Children.Add(earth);
            sun.Children.Add(mars);
            sun.Children.Add(jupiter);
            sun.Children.Add(saturn);
            sun.Children.Add(uranus);

            Group sky = new Group("Sky", GroupType.ReferenceFrame, null);

            groups.Add(sun);
            groups.Add(sky);

            return groups;
        }
    }
}
