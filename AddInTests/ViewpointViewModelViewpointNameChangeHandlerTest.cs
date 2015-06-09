// <copyright file="ViewpointViewModelViewpointNameChangeHandlerTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using Microsoft.Research.Wwt.Excel.Addin;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{
    /// <summary>
    /// This is a test class for ViewpointViewModelViewpointNameChangeHandlerTest and is intended
    /// to contain all ViewpointViewModelViewpointNameChangeHandlerTest Unit Tests
    /// </summary>
    [TestClass()]
    public class ViewpointViewModelViewpointNameChangeHandlerTest
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
        /// A test for ViewpointNameChangeHandler Constructor
        /// </summary>
        [TestMethod()]
        public void ViewpointViewModelViewpointNameChangeHandlerConstructorTest()
        {
            ViewpointViewModel viewpointViewModel = new ViewpointViewModel(ViewpointViewModelTest.GetPerspectiveInstance());
            ViewpointViewModel_Accessor.ViewpointNameChangeHandler target = new ViewpointViewModel_Accessor.ViewpointNameChangeHandler(viewpointViewModel);
            Assert.IsNotNull(viewpointViewModel);
            Assert.IsNotNull(target);
            Assert.IsNotNull(target.parent);
        }

        /// <summary>
        /// A test for Execute
        /// </summary>
        [TestMethod()]
        public void ExecuteEmptyViewpointNameTest()
        {
            ViewpointViewModel viewpointViewModel = new ViewpointViewModel(ViewpointViewModelTest.GetPerspectiveInstance());
            viewpointViewModel.PropertyChanged += new System.ComponentModel.PropertyChangedEventHandler(ViewpointViewModelPropertyChanged);

            ViewpointViewModel_Accessor.ViewpointNameChangeHandler target = new ViewpointViewModel_Accessor.ViewpointNameChangeHandler(viewpointViewModel);

            viewpointViewModel.Name = "NameBeforeEvent";

            string viewpointName = string.Empty;
            target.Execute(viewpointName);
            Assert.AreEqual(viewpointViewModel.Name, "NameAfterEvent");
        }

        /// <summary>
        /// A test for Execute
        /// </summary>
        [TestMethod()]
        public void ExecuteValidViewpointNameTest()
        {
            ViewpointViewModel viewpointViewModel = new ViewpointViewModel(ViewpointViewModelTest.GetPerspectiveInstance());
            viewpointViewModel.PropertyChanged += new System.ComponentModel.PropertyChangedEventHandler(ViewpointViewModelPropertyChanged);

            ViewpointViewModel_Accessor.ViewpointNameChangeHandler target = new ViewpointViewModel_Accessor.ViewpointNameChangeHandler(viewpointViewModel);

            viewpointViewModel.Name = "NameBeforeEvent";

            // This will be used to set the ViewpointViewModel name.
            string viewpointName = "NameAfterEvent";
            target.Execute(viewpointName);
            Assert.AreEqual(viewpointViewModel.Name, "NameAfterEvent");
        }

        /// <summary>
        /// Viewpoint view model property changed event.
        /// </summary>
        /// <param name="sender">Viewpoint View Model</param>
        /// <param name="e">Routed Event</param>
        private void ViewpointViewModelPropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            ViewpointViewModel viewpointViewModel = sender as ViewpointViewModel;
            Assert.AreEqual(viewpointViewModel.Name, "NameBeforeEvent");

            // Remove this event handler. Otherwise, this method will be called recursively.
            viewpointViewModel.PropertyChanged -= new System.ComponentModel.PropertyChangedEventHandler(ViewpointViewModelPropertyChanged);
            viewpointViewModel.Name = "NameAfterEvent";
        }
    }
}
