//-----------------------------------------------------------------------
// <copyright file="IWWTRequest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// WWTRequest interface which sends the request to WWT and gets response.
    /// </summary>
    public interface IWWTRequest
    {
        /// <summary>
        /// This function is used to send request to LCAPI.
        /// </summary>
        /// <param name="command">Uri of the API</param>
        /// <param name="payload">Data to be uploaded</param>
        /// <param name="consumeException">Whether to consume exception?</param>
        /// <returns>Response of the operation</returns>
        string Send(string command, string payload, bool consumeException);
    }
}
