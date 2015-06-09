//-----------------------------------------------------------------------
// <copyright file="GroupViewModelTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Collections.ObjectModel;
using System.Windows.Input;
using Microsoft.Research.Wwt.Excel.Addin;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{
    /// <summary>
    /// This is a test class for GroupViewModelTest and is intended
    /// to contain all GroupViewModelTest Unit Tests
    /// </summary>
    [TestClass()]
    public class GroupViewModelTest
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
        /// A test for GroupSelectionCommand
        /// </summary>
        [TestMethod()]
        public void GroupSelectionCommandTest()
        {
            GroupViewModel target = new GroupViewModel();
            ICommand actual = target.GroupSelectionCommand;
            Assert.IsNotNull(target);
            Assert.IsNotNull(actual);
        }

        /// <summary>
        /// A test for GroupViewModel Constructor
        /// </summary>
        [TestMethod()]
        public void GroupViewModelConstructorTest()
        {
            GroupViewModel target = new GroupViewModel();
            Assert.IsNotNull(target);
            Assert.IsNotNull(target.GroupSelectionCommand);
            Assert.IsNotNull(target.ReferenceGroup);
        }

        /// <summary>
        /// A test for Name
        /// </summary>
        [TestMethod()]
        public void NameTest()
        {
            GroupViewModel target = new GroupViewModel();
            string expected = "GroupViewModelName";
            target.Name = expected;
            string actual = target.Name;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for ReferenceGroup
        /// </summary>
        [TestMethod()]
        public void ReferenceGroupTest()
        {
            GroupViewModel groupViewModel = new GroupViewModel();
            Collection<Group> expected = groupViewModel.ReferenceGroup;
            Assert.IsNotNull(expected);
        }

        /// <summary>
        /// A test for ReferenceGroup
        /// </summary>
        [TestMethod()]
        public void ReferenceGroupPrivateSetTest()
        {
            GroupViewModel_Accessor target = new GroupViewModel_Accessor();
            Collection<Group> expected = new Collection<Group>();
            target.ReferenceGroup = expected;
            Collection<Group> actual = target.ReferenceGroup;
            Assert.AreEqual(expected, actual);
        }
    }
}