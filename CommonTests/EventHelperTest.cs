//-----------------------------------------------------------------------
// <copyright file="EventHelperTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// This is a test class for EventHelperTest and is intended to contain all EventHelperTest Unit Tests
    /// </summary>
    [TestClass()]
    public class EventHelperTest
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
        /// A test for OnFire
        /// </summary>
        [TestMethod()]
        public void OnFireTest()
        {
            int testValue = 0;
            EventHandler handler = delegate(object sender, EventArgs eventArgs)
            {
                if (sender.GetType() == this.GetType())
                {
                    testValue = 1;
                }
            };
            object sentFrom = this;
            EventArgs args = new EventArgs();
            EventHelper.OnFire(handler, sentFrom, args);
            int expectedTestValue = 1;
            Assert.AreEqual(expectedTestValue, testValue);
        }

        /// <summary>
        /// Helper method for OnFireGenericTest
        /// </summary>
        public void OnFireTest1Helper<TArgs>(TArgs args)
            where TArgs : EventState, new()
        {
            int testValue = 0;
            EventHandler<TArgs> handler = delegate(object sender, TArgs eventArgs)
            {
                if (sender.GetType() == this.GetType() && eventArgs.EventFired)
                {
                    testValue = 1;
                }
            };
            object sentFrom = this;
            EventHelper.OnFire<TArgs>(handler, sentFrom, args);
            int expectedTestValue = 1;
            Assert.AreEqual(expectedTestValue, testValue);
        }

        /// <summary>
        /// A test for OnFire with generic event args
        /// </summary>
        [TestMethod()]
        public void OnFireGenericTest()
        {
            EventState args = new EventState() { EventFired = true };
            this.OnFireTest1Helper(args);
        }
    }
}