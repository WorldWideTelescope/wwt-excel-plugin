//-----------------------------------------------------------------------
// <copyright file="WWTInvalidMockRequest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{
    /// <summary>
    /// Class having implementation for WWTRequest which mocks WWT and sends invalid response similar to what WWT returns.
    /// </summary>
    internal class WWTInvalidMockRequest : IWWTRequest
    {
        /// <summary>
        /// Valid machine check command.
        /// </summary>
        private const string ValidMachineCommand = @"cmd=version";

        /// <summary>
        /// Set property command
        /// </summary>
        private const string UpdateLayerCommand = @"cmd=setprops";

        /// <summary>
        /// Get all groups command.
        /// </summary>
        private const string AllGroupsCommand = @"cmd=layerlist";

        /// <summary>
        /// Update data command
        /// </summary>
        private const string UpdateDataCommand = @"cmd=update";

        /// <summary>
        /// WWT reference frame 
        /// </summary>
        private const string WWTRefernceFrames = "<?xml version=\"1.0\" encoding=\'UTF-8\'?><LayerApi><Status>Success</Status><LayerList><ReferenceFrame Name=\"Sun\" Enabled=\"True\"><ReferenceFrame Name=\"Mercury\" Enabled=\"True\" /><ReferenceFrame Name=\"Venus\" Enabled=\"True\" /> <ReferenceFrame Name=\"Earth\" Enabled=\"True\"><ReferenceFrame Name=\"Moon\" Enabled=\"True\" /><ReferenceFrame Name=\"Mars\" Enabled=\"True\"><ReferenceFrame Name=\"Phobos\" Enabled=\"True\" /><ReferenceFrame Name=\"Deimos\" Enabled=\"True\" /></ReferenceFrame><ReferenceFrame Name=\"Jupiter\" Enabled=\"True\"></ReferenceFrame><ReferenceFrame Name=\"Saturn\" Enabled=\"True\"></ReferenceFrame></ReferenceFrame></ReferenceFrame><ReferenceFrame Name=\"Sky\" Enabled=\"True\" /></LayerList></LayerApi>";

        /// <summary>
        /// Error message  which says WWT is not running.
        /// </summary>
        private const string WWTNotOpenFailure = "WorldWide Telescope (WWT) needs to be open to perform this operation. Please open WWT and try again.";

        /// <summary>
        /// WWT Success response
        /// </summary>
        private const string WWTErrorResponse = "<?xml version=\"1.0\" encoding=\"utf-8\"?><LayerApi><Status>Error - Invalid layer ID</Status></LayerApi>";

        /// <summary>
        /// Initializes a new instance of the WWTInvalidMockRequest class.
        /// </summary>
        internal WWTInvalidMockRequest()
        {
            Globals_Accessor.wwtManager = null;
        }

        /// <summary>
        /// This function is used to send request to LCAPI.
        /// </summary>
        /// <param name="command">Uri of the API</param>
        /// <param name="payload">Data to be uploaded</param>
        /// <param name="consumeException">Whether to consume exception?</param>
        /// <returns>Response of the operation</returns>
        public string Send(string command, string payload, bool consumeException)
        {
            string currentMachineUrl = string.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                    "http://{0}:5050/layerApi.aspx",
                    TargetMachine.DefaultIP);

            string response = string.Empty;

            if (!string.IsNullOrWhiteSpace(command))
            {
                if (command.Contains(ValidMachineCommand))
                {
                    response = WWTInvalidMockRequest.ProcessValidMachineCommand(command, currentMachineUrl, consumeException);
                }
                else if (command.Contains(UpdateLayerCommand) || command.Contains(UpdateDataCommand))
                {
                    response = Constants.DefaultErrorResponse;
                }
                else if (command.Contains(AllGroupsCommand))
                {
                    response = WWTRefernceFrames;
                }
            }

            return response;
        }

        /// <summary>
        /// Processes the Valid Machine command passed and gets the appropriate response.
        /// </summary>
        /// <param name="command">Uri of the API</param>
        /// <param name="currentMachineUrl">Uri of the API for the current machine</param>
        /// <param name="consumeException">Whether to consume exception?</param>
        /// <returns>Response of the Valid check command</returns>
        private static string ProcessValidMachineCommand(string command, string currentMachineUrl, bool consumeException)
        {
            string response = string.Empty;

            if (command.StartsWith(currentMachineUrl, System.StringComparison.OrdinalIgnoreCase))
            {
                response = "<?xml version=\"1.0\" encoding=\"utf-8\"?><LayerApi><Version>3.0.57</Version></LayerApi>";
            }
            else
            {
                // This will be used for negative scenario test cases.
                response = Constants.DefaultErrorResponse;
                if (!consumeException)
                {
                    throw new CustomException(WWTInvalidMockRequest.WWTNotOpenFailure, null, true, ErrorCodes.Code100001);
                }
            }

            return response;
        }
    }
}
