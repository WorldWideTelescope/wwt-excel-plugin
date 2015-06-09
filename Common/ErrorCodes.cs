//-----------------------------------------------------------------------
// <copyright file="ErrorCodes.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace Microsoft.Research.Wwt.Excel.Common
{
    public enum ErrorCodes
    {
        /// <summary>
        /// This is the default error code.
        /// </summary>
        Code000000,

        #region WWT Related codes

        /// <summary>
        /// This will be used for all WWT manger related errors.
        /// </summary>
        Code100000,

        /// <summary>
        /// WWT Not Running.
        /// </summary>
        Code100001,

        /// <summary>
        /// Could not parse Response.
        /// </summary>
        Code100002,

        /// <summary>
        /// Error response.
        /// </summary>
        Code100003,

        /// <summary>
        /// WWT Not Installed
        /// </summary>
        Code100004,

        /// <summary>
        /// Authorization to WWT fails
        /// </summary>
        Code100005,

        #endregion

        #region Excel Related codes

        /// <summary>
        /// This will be used for all Excel manger related errors.
        /// </summary>
        Code200000,

        #endregion

        #region Workflow Related codes

        /// <summary>
        /// This will be used for all workflow controller related errors.
        /// </summary>
        Code300000
        #endregion
    }
}
