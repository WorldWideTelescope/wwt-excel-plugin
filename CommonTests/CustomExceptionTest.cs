//-----------------------------------------------------------------------
// <copyright file="CustomExceptionTest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Runtime.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// This is a test class for CustomExceptionTest and is intended
    /// to contain all CustomExceptionTest Unit Tests
    /// </summary>
    [TestClass()]
    public class CustomExceptionTest
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
        /// A test for CustomException Constructor
        /// </summary>
        [TestMethod()]
        public void CustomExceptionConstructorTest()
        {
            string message = string.Empty;
            Exception innerException = null;
            bool hasCustomMessage = false;
            CustomException target = new CustomException(message, innerException, hasCustomMessage);
            bool actual = target.HasCustomMessage;
            Assert.AreEqual(hasCustomMessage, actual);
        }

        /// <summary>
        /// A test for CustomException Constructor for checking not null
        /// </summary>
        [TestMethod()]
        public void CustomExceptionConstructorNotNullTest()
        {
            Exception innerException = new ArgumentNullException(string.Empty);
            string message = innerException.Message;
            CustomException target = new CustomException(message, innerException);
            Assert.IsNotNull(target);
        }

        /// <summary>
        /// A test for CustomException Constructor to check custom message
        /// </summary>
        [TestMethod()]
        public void CustomExceptionConstructorCustomMessageTest()
        {
            string message = string.Empty;
            Exception innerException = null;
            bool hasCustomMessage = true;
            ErrorCodes errorCode = ErrorCodes.Code100004;
            CustomException target = new CustomException(message, innerException, hasCustomMessage, errorCode);
            bool actual = target.HasCustomMessage;
            Assert.AreEqual(hasCustomMessage, actual);
        }

        /// <summary>
        /// A test for CustomException Constructor to check error code
        /// </summary>
        [TestMethod()]
        public void CustomExceptionConstructorErrorCodeTest()
        {
            string message = string.Empty;
            bool hasCustomMessage = false;
            ErrorCodes errorCode = ErrorCodes.Code100001;
            CustomException target = new CustomException(message, hasCustomMessage, errorCode);
            ErrorCodes actual = target.ErrorCode;
            Assert.AreEqual(errorCode, actual);
        }

        /// <summary>
        /// A test for CustomException Constructor to check exception message
        /// </summary>
        [TestMethod()]
        public void CustomExceptionConstructorExceptionMessageTest()
        {
            string message = "Exception Message";
            CustomException target = new CustomException(message);
            string actual = target.Message;
            Assert.AreEqual(message, actual);
        }

        /// <summary>
        /// A test for CustomException Constructor
        /// </summary>
        [TestMethod()]
        public void CustomExceptionConstructorConstructorTest()
        {
            CustomException target = new CustomException();
            Assert.IsNotNull(target);
        }

        /// <summary>
        /// A test for CustomException Constructor
        /// </summary>
        [TestMethod()]
        public void CustomExceptionConstructorEmptyMessageTest()
        {
            string message = string.Empty; 
            bool hasCustomMessage = false;
            CustomException target = new CustomException(message, hasCustomMessage);
            string actual = target.Message;
            Assert.AreEqual(message, actual);
        }

        /// <summary>
        /// A test for CustomException Constructor to check error code
        /// </summary>
        [TestMethod()]
        public void CustomExceptionConstructorErrorCodeEmptyMessageTest()
        {
            string message = string.Empty;
            ErrorCodes errorCode = ErrorCodes.Code300000;
            CustomException target = new CustomException(message, errorCode);
            ErrorCodes actual = target.ErrorCode;
            Assert.AreEqual(errorCode, actual);
        }

        /// <summary>
        /// A test for HResult
        /// </summary>
        [TestMethod()]
        public void HResultTest()
        {
            CustomException target = new CustomException();
            int expected = 0;
            int actual = target.HResult;
            Assert.AreNotEqual(expected, actual);
        }

        /// <summary>
        /// A test for CustomException Constructor
        /// </summary>
        [TestMethod()]
        public void CustomExceptionConstructorTest1()
        {
            SerializationInfo info = new SerializationInfo(typeof(CustomException), new FormatterConverter());
            info.AddValue("ClassName", string.Empty);
            info.AddValue("Message", string.Empty);
            info.AddValue("InnerException", new ArgumentException(string.Empty));
            info.AddValue("HelpURL", string.Empty);
            info.AddValue("StackTraceString", string.Empty);
            info.AddValue("RemoteStackTraceString", string.Empty);
            info.AddValue("RemoteStackIndex", 0);
            info.AddValue("ExceptionMethod", string.Empty);
            info.AddValue("HResult", 1);
            info.AddValue("Source", string.Empty);

            StreamingContext context = new StreamingContext();
            CustomException_Accessor target = new CustomException_Accessor(info, context);
            Assert.IsNotNull(target);
        }

        /// <summary>
        /// A test for GetObjectData
        /// </summary>
        [TestMethod()]
        public void GetObjectDataTest()
        {
            CustomException target = new CustomException();
            SerializationInfo info = new SerializationInfo(typeof(CustomException), new FormatterConverter());
            StreamingContext context = new StreamingContext();
            target.GetObjectData(info, context);

            Assert.IsNotNull(info);
            Assert.IsNotNull(context);
        }
    }
}
