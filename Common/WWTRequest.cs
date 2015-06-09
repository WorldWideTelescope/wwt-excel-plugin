//-----------------------------------------------------------------------
// <copyright file="WWTRequest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Net;
using System.Text.RegularExpressions;
using System.Xml;

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Class having implementation for WWTRequest which talks to WWT to send the commands from Excel Add-In.
    /// </summary>
    public class WWTRequest : IWWTRequest
    {
        /// <summary>
        /// This function is used to send request to LCAPI.
        /// </summary>
        /// <param name="command">Uri of the API</param>
        /// <param name="payload">Data to be uploaded</param>
        /// <param name="consumeException">Whether to consume exception?</param>
        /// <returns>Response of the operation</returns>
        public string Send(string command, string payload, bool consumeException)
        {
            string response = string.Empty;
            using (WebClient client = new WebClient())
            {
                try
                {
                    response = client.UploadString(command, payload);
                    if (!string.IsNullOrEmpty(response))
                    {
                        XmlDocument doc = new XmlDocument();
                        doc.LoadXml(response);
                        XmlNode node = doc[Constants.LCAPIElementName];
                        string s = node.InnerText;

                        // This is valid response with error string for error happened because of the data
                        // Consuming it for the time being
                        if (s.Contains(Constants.LCAPIErrorText))
                        {
                            response = Constants.DefaultErrorResponse;
                            if (!consumeException)
                            {
                                if (s.Contains(Constants.LCAPIConnectionErrorText))
                                {
                                    // TargetMachine object is initialized to get the IP of the local machine.
                                    TargetMachine localMachine = new TargetMachine();
                                    throw new CustomException(string.Format(System.Globalization.CultureInfo.InvariantCulture, Properties.Resources.ErrorLCAPIConnectionFailure, localMachine.MachineIP), true, ErrorCodes.Code100005);
                                }
                                else
                                {
                                    throw new CustomException(Properties.Resources.ErrorFromLCAPICall, true, ErrorCodes.Code100003);
                                }
                            }
                        }
                        else
                        {
                            // sanitize the response for hex characters
                            response = Regex.Replace(response, Constants.HexCharacterPattern, string.Empty);
                        }
                    }
                }
                catch (XmlException exception)
                {
                    Logger.LogException(exception);
                    response = Constants.DefaultErrorResponse;
                    if (!consumeException)
                    {
                        throw new CustomException(Properties.Resources.ErrorFromLCAPICall, exception, true, ErrorCodes.Code100002);
                    }
                }
                catch (WebException exception)
                {
                    Logger.LogException(exception);
                    response = Constants.DefaultErrorResponse;
                    if (!consumeException)
                    {
                        throw new CustomException(Properties.Resources.WWTNotOpenFailure, exception, true, ErrorCodes.Code100001);
                    }
                }
            }

            return response;
        }
    }
}
