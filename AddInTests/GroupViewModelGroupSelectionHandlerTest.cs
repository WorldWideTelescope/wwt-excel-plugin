//-----------------------------------------------------------------------
// <copyright file="GroupViewModelGroupSelectionHandlerTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using Microsoft.Research.Wwt.Excel.Addin;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{
    /// <summary>
    /// This is a test class for GroupViewModelGroupSelectionHandlerTest and is intended
    /// to contain all GroupViewModelGroupSelectionHandlerTest Unit Tests
    /// </summary>
    [TestClass()]
    public class GroupViewModelGroupSelectionHandlerTest
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
        /// A test for GroupSelectionHandler Constructor
        /// </summary>
        [TestMethod()]
        public void GroupViewModelGroupSelectionHandlerConstructorTest()
        {
            GroupViewModel groupViewModel = new GroupViewModel();
            GroupViewModel_Accessor.GroupSelectionHandler target = new GroupViewModel_Accessor.GroupSelectionHandler(groupViewModel);
            Assert.AreEqual(target.parent, groupViewModel);
        }

        /// <summary>
        /// A test for Execute
        /// </summary>
        [TestMethod()]
        public void ExecuteTest()
        {
            GroupViewModel groupViewModel = new GroupViewModel();
            groupViewModel.GroupSelectionChangedEvent += new EventHandler(GoupViewModelGroupSelectionChangedEvent);

            groupViewModel.Name = "NameBeforeEvent";

            GroupViewModel_Accessor.GroupSelectionHandler target = new GroupViewModel_Accessor.GroupSelectionHandler(groupViewModel);
            target.Execute(groupViewModel);
            Assert.AreEqual(groupViewModel.Name, "NameAfterEvent");
        }

        /// <summary>
        /// Group selection changed event.
        /// </summary>
        /// <param name="sender">Group view model object</param>
        /// <param name="e">Routed event</param>
        private void GoupViewModelGroupSelectionChangedEvent(object sender, EventArgs e)
        {
            GroupViewModel groupViewModel = sender as GroupViewModel;
            Assert.AreEqual(groupViewModel.Name, "NameBeforeEvent");
            groupViewModel.Name = "NameAfterEvent";
        }
    }
}