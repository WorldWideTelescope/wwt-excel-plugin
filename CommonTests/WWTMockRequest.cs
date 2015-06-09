//-----------------------------------------------------------------------
// <copyright file="WWTMockRequest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// Class having implementation for WWTRequest which mocks WWT and sends response similar to what WWT returns.
    /// </summary>
    internal class WWTMockRequest : IWWTRequest 
    {
        /// <summary>
        /// Name of layer to be used in test cases
        /// </summary>
        private const string TestLayerName = @"TestLayer";

        /// <summary>
        /// Layer Id for the Test layer
        /// </summary>
        private const string TestLayerID = @"3f0cfda2-7319-4190-8f5e-99778a04ca3d";

        /// <summary>
        /// Valid machine check command.
        /// </summary>
        private const string ValidMachineCommand = @"cmd=version";

        /// <summary>
        /// Get all layers command
        /// </summary>
        private const string GetAllLayersCommand = @"cmd=layerlist";

        /// <summary>
        /// Upload data command
        /// </summary>
        private const string UpdateCommand = @"cmd=update";

        /// <summary>
        /// New layer command
        /// </summary>
        private const string NewLayerCommand = @"cmd=new";

        /// <summary>
        /// Layer name attribute
        /// </summary>
        private const string LayerNameAttribute = @"name";

        /// <summary>
        /// Frame name attribute
        /// </summary>
        private const string FrameNameAttribute = @"frame";

        /// <summary>
        /// Layer Id attribute
        /// </summary>
        private const string LayerIdAttribute = @"id";
        
        /// <summary>
        /// nopurge attribute set to true
        /// </summary>
        private const string NoPurgeTrueAttribute = @"nopurge=true";
        
        /// <summary>
        /// hasheader attribute set to true
        /// </summary>
        private const string HasHeaderTrueAttribute = @"hasheader=true";

        /// <summary>
        /// purgeall attribute set to true
        /// </summary>
        private const string PurgeAllTrueAttribute = @"purgeall=true";
       
        /// <summary>
        /// IP which will be used to send invalid version exception.
        /// </summary>
        private const string InvalidVersionIp = @"http://0.0.0.0:5050";

        /// <summary>
        /// Error message  which says WWT is not running.
        /// </summary>
        private const string WWTNotOpenFailure = "WorldWide Telescope (WWT) needs to be open to perform this operation. Please open WWT and try again.";

        /// <summary>
        /// Initializes a new instance of the WWTMockRequest class.
        /// </summary>
        internal WWTMockRequest()
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
                    response = WWTMockRequest.ProcessValidMachineCommand(command, currentMachineUrl, consumeException);
                }
                else if (command.Contains(GetAllLayersCommand))
                {
                    response = WWTMockRequest.ProcessGetAllLayersCommand(command, currentMachineUrl, consumeException);
                }
                else if (command.Contains(UpdateCommand) && command.Contains(NoPurgeTrueAttribute) && command.Contains(HasHeaderTrueAttribute))
                {
                    response = WWTMockRequest.ProcessUploadDataCommand(command, currentMachineUrl, consumeException);
                }
                else if (command.Contains(UpdateCommand) && command.Contains(NoPurgeTrueAttribute) && command.Contains(PurgeAllTrueAttribute))
                {
                    response = WWTMockRequest.ProcessPurgeDataCommand(command, currentMachineUrl, consumeException);
                }
                else if (command.Contains(NewLayerCommand) && command.Contains(LayerNameAttribute) && command.Contains(FrameNameAttribute))
                {
                    response = WWTMockRequest.ProcessCreateLayerCommand(command, currentMachineUrl, consumeException);
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
                    if (command.StartsWith(WWTMockRequest.InvalidVersionIp, System.StringComparison.OrdinalIgnoreCase))
                    {
                        throw new CustomException(WWTMockRequest.WWTNotOpenFailure, null, true, ErrorCodes.Code100002);
                    }
                    else
                    {
                        throw new CustomException(WWTMockRequest.WWTNotOpenFailure, null, true, ErrorCodes.Code100001);
                    }
                }
            }

            return response;
        }

        /// <summary>
        /// Processes the GetAllLayers command passed and gets the appropriate response.
        /// </summary>
        /// <param name="command">Uri of the API</param>
        /// <param name="currentMachineUrl">Uri of the API for the current machine</param>
        /// <param name="consumeException">Whether to consume exception?</param>
        /// <returns>Response of the command</returns>
        private static string ProcessGetAllLayersCommand(string command, string currentMachineUrl, bool consumeException)
        {
            string response = string.Empty;

            if (command.StartsWith(currentMachineUrl, System.StringComparison.OrdinalIgnoreCase))
            {
                response = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><LayerApi><Status>Success</Status><LayerList><ReferenceFrame Name=\"Earth\" Enabled=\"True\"><Layer Name=\"TestLayer_1\" ID=\"c71ebb83-a2b2-437b-8cef-1524b4c8aa7e\" Type=\"SpreadSheetLayer\" Enabled=\"True\" /><ReferenceFrame Name=\"Moon\" Enabled=\"True\"><Layer Name=\"TestLayer_2\" ID=\"166e7aa9-c84d-4e81-b0d7-57467a13b601\" Type=\"SpreadSheetLayer\" Enabled=\"True\" /></ReferenceFrame></ReferenceFrame><ReferenceFrame Name=\"Sky\" Enabled=\"True\"/></LayerList></LayerApi>";
            }
            else
            {
                // This will be used for negative scenario test cases.
                response = Constants.DefaultErrorResponse;
                if (!consumeException)
                {
                    throw new CustomException(WWTMockRequest.WWTNotOpenFailure, null, true, ErrorCodes.Code100001);
                }
            }

            return response;
        }

        /// <summary>
        /// Processes the Upload Data command passed and gets the appropriate response.
        /// </summary>
        /// <param name="command">Uri of the API</param>
        /// <param name="currentMachineUrl">Uri of the API for the current machine</param>
        /// <param name="consumeException">Whether to consume exception?</param>
        /// <returns>Response of the command</returns>
        private static string ProcessUploadDataCommand(string command, string currentMachineUrl, bool consumeException)
        {
            string response = string.Empty;

            if (command.StartsWith(currentMachineUrl, System.StringComparison.OrdinalIgnoreCase))
            {
                response = "<?xml version=\"1.0\" encoding=\"utf-8\"?><LayerApi><Status>Success</Status></LayerApi>";
            }
            else
            {
                // This will be used for negative scenario test cases.
                response = Constants.DefaultErrorResponse;
                if (!consumeException)
                {
                    throw new CustomException(WWTMockRequest.WWTNotOpenFailure, null, true, ErrorCodes.Code100001);
                }
            }

            return response;
        }

        /// <summary>
        /// Processes the Purge Data command passed and gets the appropriate response.
        /// </summary>
        /// <param name="command">Uri of the API</param>
        /// <param name="currentMachineUrl">Uri of the API for the current machine</param>
        /// <param name="consumeException">Whether to consume exception?</param>
        /// <returns>Response of the command</returns>
        private static string ProcessPurgeDataCommand(string command, string currentMachineUrl, bool consumeException)
        {
            string response = string.Empty;

            if (command.StartsWith(currentMachineUrl, System.StringComparison.OrdinalIgnoreCase) && ParseCommand(command, LayerIdAttribute) == TestLayerID)
            {
                response = "<?xml version=\"1.0\" encoding=\"utf-8\"?><LayerApi><Status>Success</Status></LayerApi>";
            }
            else
            {
                // This will be used for negative scenario test cases.
                response = Constants.DefaultErrorResponse;
                if (!consumeException)
                {
                    throw new CustomException(WWTMockRequest.WWTNotOpenFailure, null, true, ErrorCodes.Code100001);
                }
            }

            return response;
        }

        /// <summary>
        /// Processes the Create Layer command passed and gets the appropriate response.
        /// </summary>
        /// <param name="command">Uri of the API</param>
        /// <param name="currentMachineUrl">Uri of the API for the current machine</param>
        /// <param name="consumeException">Whether to consume exception?</param>
        /// <returns>Response of the command</returns>
        private static string ProcessCreateLayerCommand(string command, string currentMachineUrl, bool consumeException)
        {
            string response = string.Empty;

            if (command.StartsWith(currentMachineUrl, System.StringComparison.OrdinalIgnoreCase) && ParseCommand(command, LayerNameAttribute) == TestLayerName)
            {
                response = "<?xml version=\"1.0\" encoding=\"utf-8\"?><LayerApi><NewLayerID>" + TestLayerID + "</NewLayerID></LayerApi>";
            }
            else
            {
                // This will be used for negative scenario test cases.
                response = Constants.DefaultErrorResponse;
                if (!consumeException)
                {
                    throw new CustomException(WWTMockRequest.WWTNotOpenFailure, null, true, ErrorCodes.Code100001);
                }
            }

            return response;
        }

        private static string ParseCommand(string command, string attributeName)
        {
            string value = string.Empty;
            string[] cmdPartsOne = command.Split('&');
            foreach (string part in cmdPartsOne)
            {
                if (part.Contains(attributeName + "="))
                {
                    string[] cmdPartsTwo = part.Split('=');
                    value = cmdPartsTwo[1];
                }
            }

            return value;
        }
    }
}
