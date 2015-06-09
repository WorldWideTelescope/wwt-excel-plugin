//-----------------------------------------------------------------------
// <copyright file="WWTMockRequest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.AddIn.Tests
{
    /// <summary>
    /// Class having implementation for WWTRequest which mocks WWT and sends response similar to what WWT returns.
    /// </summary>
    internal class WWTMockRequest : IWWTRequest
    {
        /// <summary>
        /// Valid machine check command.
        /// </summary>
        private const string ValidMachineCommand = @"cmd=version";

        /// <summary>
        /// Get all groups command.
        /// </summary>
        private const string AllGroupsCommand = @"cmd=layerlist";

        /// <summary>
        /// Set property command
        /// </summary>
        private const string UpdateLayerCommand = @"cmd=setprops";

        /// <summary>
        /// Get Layer header command
        /// </summary>
        private const string GetLayerHeaderCommand = @"cmd=get";

        /// <summary>
        /// Create layer group command
        /// </summary>
        private const string CreateLayerGroupCommand = @"cmd=group&frame";

        /// <summary>
        /// Create layer command
        /// </summary>
        private const string CreateLayerCommand = @"cmd=new&name";

        /// <summary>
        /// Update data command
        /// </summary>
        private const string UpdateDataCommand = @"cmd=update";

        /// <summary>
        /// Get camera view command.
        /// </summary>
        private const string GetCameraViewCommand = @"cmd=state";

        /// <summary>
        /// Get Layer details command
        /// </summary>
        private const string GetLayerDetailsCommand = @"cmd=getprops&id=";

        /// <summary>
        /// Set Mode Command
        /// </summary>
        private const string SetModeCommand = @"cmd=mode&lookat";

        /// <summary>
        /// Set Camera View Command
        /// </summary>
        private const string SetCameraViewCommand = @"cmd=state&instant=false";

        /// <summary>
        /// Error message  which says WWT is not running.
        /// </summary>
        private const string WWTNotOpenFailure = "WorldWide Telescope (WWT) needs to be open to perform this operation. Please open WWT and try again.";

        /// <summary>
        /// WWT reference frame 
        /// </summary>
        private const string WWTRefernceFrames = "<?xml version=\"1.0\" encoding=\'UTF-8\'?><LayerApi><Status>Success</Status><LayerList><ReferenceFrame Name=\"Sun\" Enabled=\"True\"><ReferenceFrame Name=\"Mercury\" Enabled=\"True\" /><ReferenceFrame Name=\"Venus\" Enabled=\"True\" /> <ReferenceFrame Name=\"Earth\" Enabled=\"True\"><ReferenceFrame Name=\"Moon\" Enabled=\"True\" /><ReferenceFrame Name=\"Mars\" Enabled=\"True\"><ReferenceFrame Name=\"Phobos\" Enabled=\"True\" /><ReferenceFrame Name=\"Deimos\" Enabled=\"True\" /></ReferenceFrame><ReferenceFrame Name=\"Jupiter\" Enabled=\"True\"></ReferenceFrame><ReferenceFrame Name=\"Saturn\" Enabled=\"True\"></ReferenceFrame></ReferenceFrame></ReferenceFrame><ReferenceFrame Name=\"Sky\" Enabled=\"True\" /></LayerList></LayerApi>";

        /// <summary>
        /// WWT Success 
        /// </summary>
        private const string WWTSuccessResponse = "<?xml version=\"1.0\" encoding=\"utf-8\"?><LayerApi><Status>Success</Status></LayerApi>";

        /// <summary>
        /// Create layer response
        /// </summary>
        private const string WWTCreateResponse = "<?xml version=\"1.0\" encoding=\"utf-8\"?><LayerApi><NewLayerID>2cf4374f-e1ce-47a9-b08c-31079765ddcf</NewLayerID></LayerApi>";

        /// <summary>
        /// Get layer details response
        /// </summary>
        private const string WWTGetlayerDetailsResponse = "<?xml version='1.0' encoding='UTF-8'?><LayerApi><Status>Success</Status><Layer Class=\"SpreadSheetLayer\" BeginRange=\"1/9/2009 3:44:38 AM\" EndRange=\"1/22/2009 2:34:29 PM\" Decay=\"16\" CoordinatesType=\"Spherical\" LatColumn=\"0\" LngColumn=\"1\" GeometryColumn=\"-1\" XAxisColumn=\"-1\" YAxisColumn=\"-1\" ZAxisColumn=\"-1\" XAxisReverse=\"False\" YAxisReverse=\"False\" ZAxisReverse=\"False\" AltType=\"Depth\" MarkerMix=\"Same_For_All\" RaUnits=\"Hours\" MarkerColumn=\"-1\" ColorMapColumn=\"6\" PlotType=\"Gaussian\" MarkerIndex=\"0\" ShowFarSide=\"False\" MarkerScale=\"World\" AltUnit=\"Kilometers\" CartesianScale=\"Meters\" CartesianCustomScale=\"1\" AltColumn=\"2\" StartDateColumn=\"3\" EndDateColumn=\"-1\" SizeColumn=\"7\" NameColumn=\"0\" HyperlinkFormat=\"\" HyperlinkColumn=\"-1\" ScaleFactor=\"1\" PointScaleType=\"Power\" Opacity=\"1\" StartTime=\"1/1/0001 12:00:00 AM\" EndTime=\"12/31/9999 11:59:59 PM\" FadeSpan=\"00:00:00\" FadeType=\"None\" Name=\"NewLayer\" ColorValue=\"ARGBColor:255:255:255:255\" Enabled=\"True\" Astronomical=\"False\" /></LayerApi>";

        /// <summary>
        /// WWT get camera view response.
        /// </summary>
        private const string WWTGetCameraViewResponse = "<?xml version=\"1.0\" encoding=\"utf-8\"?><LayerApi><Status>Success</Status><ViewState lookat=\"Earth\" lat=\"0\" lng=\"0\" zoom=\"360\" angle=\"0\" rotation=\"0\" time=\"3/8/2011 5:02:14 PM\" timerate=\"1\" ReferenceFrame=\"Earth\" ViewToken=\"SD8834DFA\" ZoomText=\"59200 km\"></ViewState></LayerApi>";

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
                else if (command.Contains(AllGroupsCommand))
                {
                    response = WWTRefernceFrames;
                }
                else if (command.Contains(UpdateLayerCommand) ||
                    command.Contains(CreateLayerGroupCommand) ||
                    command.Contains(SetModeCommand) ||
                    command.Contains(SetCameraViewCommand))
                {
                    response = WWTSuccessResponse;
                }
                else if (command.Contains(CreateLayerCommand))
                {
                    response = WWTCreateResponse;
                }
                else if (command.Contains(UpdateDataCommand))
                {
                    response = WWTSuccessResponse;
                }
                else if (command.Contains(GetCameraViewCommand))
                {
                    response = WWTGetCameraViewResponse;
                }
                else if (command.Contains(GetLayerDetailsCommand))
                {
                    response = WWTGetlayerDetailsResponse;
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
                    throw new CustomException(WWTMockRequest.WWTNotOpenFailure, null, true, ErrorCodes.Code100001);
                }
            }

            return response;
        }
    }
}
