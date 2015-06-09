//-----------------------------------------------------------------------
// <copyright file="ViewpointViewModelTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Windows.Input;
using Microsoft.Research.Wwt.Excel.Addin;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{
    /// <summary>
    /// This is a test class for ViewpointViewModelTest and is intended
    /// to contain all ViewpointViewModelTest Unit Tests
    /// </summary>
    [TestClass()]
    public class ViewpointViewModelTest
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
        /// A test for Name
        /// </summary>
        [TestMethod()]
        public void NameTest()
        {
            Perspective perspective = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(perspective);
            string expected = "ViewpointViewModelName";
            target.Name = expected;
            string actual = target.Name;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for ViewpointViewModel Constructor
        /// </summary>
        [TestMethod()]
        public void ViewpointViewModelConstructorTest()
        {
            Perspective perspective = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(perspective);
            Assert.IsNotNull(target);
            Assert.IsNotNull(target.CurrentPerspective);
            Assert.IsNotNull(target.ViewpointNameChangeCommand);
            Assert.IsNotNull(target.ViewpointNameTextChangeCommand);
        }

        /// <summary>
        /// A test for OnRequestClose
        /// </summary>
        [TestMethod()]
        public void OnRequestCloseTest()
        {
            Perspective perspective = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(perspective);
            target.Name = "NameBeforeEvent";
            target.RequestClose += new System.EventHandler(ViewpointViewModelRequestClose);
            target.OnRequestClose();
            Assert.AreEqual(target.Name, "NameAfterEvent");
        }

        /// <summary>
        /// A test for CurrentPerspective
        /// </summary>
        [TestMethod()]
        public void CurrentPerspectiveTest()
        {
            Perspective expected = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(expected);
            target.CurrentPerspective = expected;
            Perspective actual = target.CurrentPerspective;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsButtonEnabled
        /// </summary>
        [TestMethod()]
        public void IsButtonEnabledTest()
        {
            Perspective perspective = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(perspective);
            bool expected = false;
            target.IsButtonEnabled = expected;
            bool actual = target.IsButtonEnabled;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsNotSky
        /// </summary>
        [TestMethod()]
        public void IsNotSkyGetFalseTest()
        {
            Perspective perspective = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(perspective);
            bool expected = true;
            bool actual = target.IsNotSky;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsNotSky
        /// </summary>
        [TestMethod()]
        public void IsNotSkyGetTrueTest()
        {
            Perspective perspective = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(perspective);
            target.CurrentPerspective.HasRADec = true;
            bool expected = false;
            bool actual = target.IsNotSky;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsNotSky
        /// </summary>
        [TestMethod()]
        public void IsNotSkySetFalseTest()
        {
            Perspective perspective = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(perspective);
            bool expected = false;
            target.IsNotSky = false;
            bool actual = target.IsNotSky;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsNotSky
        /// </summary>
        [TestMethod()]
        public void IsNotSkySetTrueTest()
        {
            Perspective perspective = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(perspective);
            bool expected = true;
            target.IsNotSky = true;
            bool actual = target.IsNotSky;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsSelected
        /// </summary>
        [TestMethod()]
        public void IsSelectedTest()
        {
            Perspective perspective = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(perspective);
            bool expected = false;
            target.IsSelected = expected;
            bool actual = target.IsSelected;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsSky
        /// </summary>
        [TestMethod()]
        public void IsSkyGetTrueTest()
        {
            Perspective perspective = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(perspective);
            target.CurrentPerspective.HasRADec = true;
            bool expected = true;
            bool actual = target.IsSky;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsSky
        /// </summary>
        [TestMethod()]
        public void IsSkyGetFalseTest()
        {
            Perspective perspective = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(perspective);
            bool expected = false;
            bool actual = target.IsSky;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsSky
        /// </summary>
        [TestMethod()]
        public void IsSkySetTrueTest()
        {
            Perspective perspective = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(perspective);
            bool expected = false;
            target.IsSky = expected;
            bool actual = target.IsSky;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for IsSky
        /// </summary>
        [TestMethod()]
        public void IsSkySetFalseTest()
        {
            Perspective perspective = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(perspective);
            bool expected = true;
            target.IsSky = expected;
            bool actual = target.IsSky;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for ViewpointNameTextChangeCommand
        /// </summary>
        [TestMethod()]
        public void ViewpointNameTextChangeCommandTest()
        {
            Perspective perspective = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(perspective);
            ICommand actual = target.ViewpointNameTextChangeCommand;
            Assert.IsNotNull(actual);
        }

        /// <summary>
        /// A test for ViewpointNameChangeCommand
        /// </summary>
        [TestMethod()]
        public void ViewpointNameChangeCommandTest()
        {
            Perspective perspective = GetPerspectiveInstance();
            ViewpointViewModel target = new ViewpointViewModel(perspective);
            ICommand actual = target.ViewpointNameChangeCommand;
            Assert.IsNotNull(actual);
        }

        /// <summary>
        /// Gets an perspective object instance which will be used by unit test cases.
        /// </summary>
        /// <returns>Perspective object</returns>
        internal static Perspective GetPerspectiveInstance()
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
            return new Perspective(lookAt, referenceFrame, hasRADec, latitude, longitude, zoom, rotation, lookAngle, observingTime, timeRate, zoomText, viewToken);
        }

        /// <summary>
        /// OnRequestClose event handler for Viewpoint View Model
        /// </summary>
        /// <param name="sender">Viewpoint View Model</param>
        /// <param name="e">Routed Event</param>
        private void ViewpointViewModelRequestClose(object sender, System.EventArgs e)
        {
            ViewpointViewModel viewpointViewModel = sender as ViewpointViewModel;
            Assert.AreEqual(viewpointViewModel.Name, "NameBeforeEvent");
            viewpointViewModel.Name = "NameAfterEvent";
        }
    }
}
