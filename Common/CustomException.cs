//-----------------------------------------------------------------------
// <copyright file="CustomException.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Security.Permissions;

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Class CustomException Imitate Exception Class 
    /// </summary>
    [Serializable]
    public class CustomException : Exception
    {
        /// <summary>
        /// Flag that indicates if the message is a custom message that needs to be shown on the UI
        /// </summary>
        private bool hasCustomMessage;

        /// <summary>
        /// Error code of the exception.
        /// </summary>
        private ErrorCodes errorCode = ErrorCodes.Code000000;

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the CustomException class.
        /// </summary>
        public CustomException()
            : base()
        {
        }

        #region With Custom Error Message

        /// <summary>
        /// Initializes a new instance of the CustomException class.
        /// </summary>
        /// <param name="message">
        /// String Message
        /// </param>
        public CustomException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the CustomException class.
        /// </summary>
        /// <param name="message">
        /// String Message
        /// </param>
        /// <param name="errorCode">
        /// Error Code.
        /// </param>
        public CustomException(string message, ErrorCodes errorCode)
            : base(message)
        {
            this.errorCode = errorCode;
        }

        /// <summary>
        /// Initializes a new instance of the CustomException class. 
        /// </summary>
        /// <param name="message">
        /// String Message
        /// </param>
        /// <param name="hasCustomMessage">
        /// true if the custom message needs to be shown on the UI
        /// </param>
        public CustomException(string message, bool hasCustomMessage)
            : base(message)
        {
            this.hasCustomMessage = hasCustomMessage;
        }

        /// <summary>
        /// Initializes a new instance of the CustomException class. 
        /// </summary>
        /// <param name="message">
        /// String Message
        /// </param>
        /// <param name="hasCustomMessage">
        /// true if the custom message needs to be shown on the UI
        /// </param>
        /// <param name="errorCode">
        /// Error Code.
        /// </param>
        public CustomException(string message, bool hasCustomMessage, ErrorCodes errorCode)
            : base(message)
        {
            this.hasCustomMessage = hasCustomMessage;
            this.errorCode = errorCode;
        }

        /// <summary>
        /// Initializes a new instance of the CustomException class. 
        /// </summary>
        /// <param name="message">
        /// Message that describes the Error
        /// </param>
        /// <param name="innerException">
        /// The Exception is the cause of Current Exception
        /// </param>
        public CustomException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        /// <summary>
        /// Initializes a new instance of the CustomException class. 
        /// </summary>
        /// <param name="message">
        /// Message that describes the Error
        /// </param>
        /// <param name="innerException">
        /// The Exception is the cause of Current Exception
        /// </param>
        /// <param name="hasCustomMessage">
        /// true if the custom message needs to be shown on the UI
        /// </param>
        public CustomException(string message, Exception innerException, bool hasCustomMessage)
            : base(message, innerException)
        {
            this.hasCustomMessage = hasCustomMessage;
        }

        /// <summary>
        /// Initializes a new instance of the CustomException class. 
        /// </summary>
        /// <param name="message">
        /// Message that describes the Error
        /// </param>
        /// <param name="innerException">
        /// The Exception is the cause of Current Exception
        /// </param>
        /// <param name="hasCustomMessage">
        /// true if the custom message needs to be shown on the UI
        /// </param>
        /// <param name="errorCode">
        /// Error Code.
        /// </param>
        public CustomException(string message, Exception innerException, bool hasCustomMessage, ErrorCodes errorCode)
            : base(message, innerException)
        {
            this.hasCustomMessage = hasCustomMessage;
            this.errorCode = errorCode;
        }

        /// <summary>
        ///  Initializes a new instance of the CustomException class. 
        /// </summary>
        /// <param name="info">
        /// Holds the serialized object data about the Exception being thrown
        /// </param>
        /// <param name="context">
        /// Instance of System.Runtime.Serialization.StreamingContext
        /// </param>
        protected CustomException(System.Runtime.Serialization.SerializationInfo info, System.Runtime.Serialization.StreamingContext context)
            : base(info, context)
        {
        }

        #endregion With Custom Error Message

        #endregion

        #region Properties

        /// <summary>
        /// Gets the coded numerical value that is assigned to a specific exception
        /// </summary>
        public new int HResult
        {
            get
            {
                return base.HResult;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the message is a custom message that needs to be shown on the UI
        /// </summary>
        public bool HasCustomMessage
        {
            get
            {
                return this.hasCustomMessage;
            }
        }

        /// <summary>
        /// Gets the error code.
        /// </summary>
        public ErrorCodes ErrorCode
        {
            get
            {
                return this.errorCode;
            }
        }
        #endregion

        /// <summary>
        /// gets System.Runtime.Serialization.SerializationInfo about the exception
        /// </summary>
        /// <param name="info">
        /// Instance of System.Runtime.Serialization.SerializationInfo
        /// </param>
        /// <param name="context">
        /// Instance of System.Runtime.Serialization.StreamingContext
        /// </param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2135:SecurityRuleSetLevel2MethodsShouldNotBeProtectedWithLinkDemandsFxCopRule", Justification = "Not needed.")]
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.SerializationFormatter)]
        public override void GetObjectData(System.Runtime.Serialization.SerializationInfo info, System.Runtime.Serialization.StreamingContext context)
        {
            base.GetObjectData(info, context);
        }
    }
}
