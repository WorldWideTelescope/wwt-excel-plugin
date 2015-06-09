//-----------------------------------------------------------------------
// <copyright file="GroupTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// This is a test class for GroupTest and is intended
    /// to contain all GroupTest Unit Tests
    /// </summary>
    [TestClass()]
    public class GroupTest
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

        #region GroupConstructorTest

        /// <summary>
        /// A test for Group Constructor for no parent.
        /// </summary>
        [TestMethod()]
        public void GroupConstructorTestNoParent()
        {
            string name = "TestName";
            GroupType type = GroupType.ReferenceFrame;
            Group parent = null;
            Group target = new Group(name, type, parent);
            Assert.IsNotNull(target);
            Assert.IsNull(target.Parent);
            Assert.AreEqual(target.Name, name);
            Assert.AreEqual(target.GroupType, type);
            Assert.AreEqual(target.Path, "/" + name);
            Assert.AreEqual(target.IsDeleted, false);
        }

        /// <summary>
        /// A test for Group Constructor for no parent.
        /// </summary>
        [TestMethod()]
        public void GroupConstructorTestWithParent()
        {
            string name = "Earth";
            GroupType type = GroupType.ReferenceFrame;
            Group parent = new Group("Sun", GroupType.ReferenceFrame, null);
            Group target = new Group(name, type, parent);
            Assert.IsNotNull(target);
            Assert.AreEqual(target.Parent, parent);
            Assert.AreEqual(target.Name, name);
            Assert.AreEqual(target.GroupType, type);
            Assert.AreEqual(target.Path, "/Sun/" + name);
            Assert.AreEqual(target.IsDeleted, false);
        }

        #endregion

        #region EqualsTest

        /// <summary>
        /// A test for Equals if the other same as the target.  
        /// </summary>
        [TestMethod()]
        public void EqualsTestSameReferenceFrame()
        {
            Group parent = new Group("Sun", GroupType.ReferenceFrame, null);
            Group target = new Group("Earth", GroupType.ReferenceFrame, parent);

            Group other = target;
            bool expected = true;
            bool actual;
            actual = target.Equals(other);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for Equals if the other is null.
        /// </summary>
        [TestMethod()]
        public void EqualsTestNull()
        {
            Group parent = new Group("Sun", GroupType.ReferenceFrame, null);
            Group target = new Group("Earth", GroupType.ReferenceFrame, parent);

            Group other = null;
            bool expected = false;
            bool actual;
            actual = target.Equals(other);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for Equals if the name of the two groups are different.
        /// </summary>
        [TestMethod()]
        public void EqualsTestDifferentReferenceFrameName()
        {
            Group parent = new Group("Sun", GroupType.ReferenceFrame, null);
            Group target = new Group("Earth", GroupType.ReferenceFrame, parent);

            Group other = new Group("Earth1", GroupType.ReferenceFrame, parent);
            bool expected = false;
            bool actual;
            actual = target.Equals(other);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for Equals if the path of the two groups are different.
        /// </summary>
        [TestMethod()]
        public void EqualsTestDifferentReferenceFramePath()
        {
            Group parent = new Group("Sun", GroupType.ReferenceFrame, null);
            Group target = new Group("Earth", GroupType.ReferenceFrame, parent);

            Group other = new Group("Earth", GroupType.ReferenceFrame, null);
            bool expected = false;
            bool actual;
            actual = target.Equals(other);
            Assert.AreEqual(expected, actual);
        }

        #endregion

        #region ToStringTest

        /// <summary>
        /// A test for ToString
        /// </summary>
        [TestMethod()]
        public void ToStringTestValid()
        {
            Group target = new Group("Sun", GroupType.ReferenceFrame, null);
            string expected = "Name = Sun , Path = /Sun";
            string actual;
            actual = target.ToString();
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for ToString
        /// </summary>
        [TestMethod()]
        public void ToStringTestValidWithParent()
        {
            Group parent = new Group("Sun", GroupType.ReferenceFrame, null);
            Group target = new Group("Earth", GroupType.ReferenceFrame, parent);
            string expected = "Name = Earth , Path = /Sun/Earth";
            string actual;
            actual = target.ToString();
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for ToString
        /// </summary>
        [TestMethod()]
        public void ToStringTestNullGroup()
        {
            Group target = new Group(string.Empty, GroupType.ReferenceFrame, null);
            string expected = "Name =  , Path = /";
            string actual;
            actual = target.ToString();
            Assert.AreEqual(expected, actual);
        }

        #endregion
    }
}
