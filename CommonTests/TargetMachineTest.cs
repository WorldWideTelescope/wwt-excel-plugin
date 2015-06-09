//-----------------------------------------------------------------------
// <copyright file="TargetMachineTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Collections.ObjectModel;
using System.Net;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// This is a test class for TargetMachineTest and is intended to contain all TargetMachineTest Unit Tests
    /// </summary>
    [TestClass()]
    public class TargetMachineTest
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
        /// A test for GetDefaultIp
        /// </summary>
        [TestMethod()]
        public void GetDefaultIpTest()
        {
            IPAddress expected = TargetMachine.DefaultIP;

            // Get the default IP.
            IPAddress actual = TargetMachine_Accessor.GetDefaultIp();

            Assert.IsNotNull(actual);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// A test for TargetMachine Constructor
        /// </summary>
        [TestMethod()]
        public void TargetMachineConstructorDefaultTest()
        {
            TargetMachine target = new TargetMachine();

            // Make sure target is not null.
            Assert.IsNotNull(target);

            // Make sure MachineIP is not null.
            Assert.IsNotNull(target.MachineIP);

            // Make sure DisplayValue is not null.
            Assert.IsNotNull(target.DisplayValue);
        }

        /// <summary>
        /// A test for TargetMachine Constructor
        /// </summary>
        [TestMethod()]
        [ExpectedException(typeof(CustomException))]
        public void TargetMachineConstructorWithInvalidIPAddressTest()
        {
            // Random IP address (Invalid)
            string inputValue = "121.242.121.242";

            // User WWTMockRequest in WWTManager so that all calls to WWT API will succeed.
            Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());

            TargetMachine target = new TargetMachine(inputValue);
            Assert.Fail("TargetMachine object initialized with invalid IP Address!");
        }

        /// <summary>
        /// A test for TargetMachine Constructor
        /// </summary>
        [TestMethod()]
        [ExpectedException(typeof(CustomException))]
        public void TargetMachineConstructorWithInvalidParameterTest()
        {
            // Invalid parameter
            string inputValue = "InvalidParameter";

            // User WWTMockRequest in WWTManager so that all calls to WWT API will succeed.
            Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());

            TargetMachine target = new TargetMachine(inputValue);
            Assert.Fail("TargetMachine object initialized with invalid parameter!");
        }

        /// <summary>
        /// A test for SetMachineIPIfValid
        /// </summary>
        [TestMethod()]
        public void SetMachineIPIfValidListTest()
        {
            // Use WWTMockRequest in WWTManager so that all calls to WWT API will succeed.
            Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());

            TargetMachine_Accessor target = new TargetMachine_Accessor();

            string machineName = Constants.Localhost;
            Collection<IPAddress> machineAddresses = TargetMachine_Accessor.GetIpFromName(machineName);
            target.SetMachineIPIfValid(machineAddresses, machineName);
            
            Assert.AreEqual(Constants.Localhost, target.DisplayValue);
            Assert.AreEqual(TargetMachine_Accessor.GetDefaultIp(), target.MachineIP);
        }

        /// <summary>
        /// A test for SetMachineIPIfValid
        /// </summary>
        [TestMethod()]
        [ExpectedException(typeof(CustomException))]
        public void SetMachineIPIfValidListExceptionTest()
        {
            // Use WWTMockRequest in WWTManager so that all calls to WWT API will succeed.
            Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());

            TargetMachine_Accessor target = new TargetMachine_Accessor();

            string machineName = Constants.Localhost;
            Collection<IPAddress> machineAddresses = new Collection<IPAddress>();
            machineAddresses.Add(new IPAddress(new byte[] { 1, 2, 3, 4 }));
            machineAddresses.Add(new IPAddress(new byte[] { 1, 2, 3, 5 }));
            target.SetMachineIPIfValid(machineAddresses, machineName);

            Assert.AreEqual(TargetMachine_Accessor.GetDefaultIp(), target.MachineIP);
        }

        /// <summary>
        /// A test for SetMachineIPIfValid
        /// </summary>
        [TestMethod()]
        [ExpectedException(typeof(CustomException))]
        public void SetMachineIPIfValidInvalidVersionExceptionTest()
        {
            // Use WWTMockRequest in WWTManager so that all calls to WWT API will succeed.
            Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());

            TargetMachine_Accessor target = new TargetMachine_Accessor();

            string machineName = Constants.Localhost;
            Collection<IPAddress> machineAddresses = new Collection<IPAddress>();
            machineAddresses.Add(new IPAddress(new byte[] { 0, 0, 0, 0 }));
            target.SetMachineIPIfValid(machineAddresses, machineName);

            Assert.AreEqual(Constants.Localhost, target.DisplayValue);
            Assert.AreEqual(TargetMachine_Accessor.GetDefaultIp(), target.MachineIP);
        }

        /// <summary>
        /// A test for SetMachineIPIfValid
        /// </summary>
        [TestMethod()]
        public void SetMachineIPIfValidTest()
        {
            // Use WWTMockRequest in WWTManager so that all calls to WWT API will succeed.
            Globals_Accessor.wwtManager = new WWTManager(new WWTMockRequest());

            TargetMachine_Accessor target = new TargetMachine_Accessor();

            IPAddress machineAddress = TargetMachine_Accessor.GetDefaultIp();
            string displayName = Constants.Localhost;
            target.SetMachineIPIfValid(machineAddress, displayName);

            Assert.AreEqual(Constants.Localhost, target.DisplayValue);
            Assert.AreEqual(TargetMachine_Accessor.GetDefaultIp(), target.MachineIP);
        }
    }
}
