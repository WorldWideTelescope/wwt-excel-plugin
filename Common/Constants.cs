//-----------------------------------------------------------------------
// <copyright file="Constants.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Constants class
    /// </summary>
    public static class Constants
    {
        #region Application Registry Properties

        /// <summary>
        /// The application registry path.
        /// </summary>
        public const string AppRegistryPath = "Software\\Microsoft\\Office\\Excel\\Addins\\Microsoft.Research.Wwt.Excel.Addin";

        /// <summary>
        /// Uninstall registry path.
        /// </summary>
        public const string UninstallRegistryPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall";

        /// <summary>
        /// The application registry key.
        /// </summary>
        public const string AppRegistryKey = "Manifest";

        /// <summary>
        /// WTML file extension used for searching WWT shell open command location in the registry
        /// </summary>
        public const string WtmlFileExtension = ".wtml";

        /// <summary>
        /// Shell open command registry key that locates WWT application path
        /// </summary>
        public const string ShellOpenCommandRegistryKey = @"\shell\open\command";

        #endregion

        #region Perspective Properties

        /// <summary>
        /// Azimuth Default Value in degrees
        /// </summary>
        public const string RotationDefaultValue = "0";

        /// <summary>
        /// Longitude Default Value in degrees
        /// </summary>
        public const string LongitudeDefaultValue = "0";

        /// <summary>
        /// Latitude Default Value in degrees
        /// </summary>
        public const string LatitudeDefaultValue = "0";

        /// <summary>
        /// LookAngle Default Value in degrees
        /// </summary>
        public const string LookAngleDefaultValue = "0";

        /// <summary>
        /// TimeRate Default Value 
        /// Set on the scale of 1 to 10000000000
        /// </summary>
        public const string TimeRateDefaultValue = "1";

        /// <summary>
        /// Zoom Default Value
        /// Set on a scale of 0 - 360 
        /// with 360 representing 59200 KM
        /// </summary>
        public const string ZoomDefaultValue = "360";

        /// <summary>
        /// Earth Zoom Text Default Value
        /// </summary>
        public const string EarthZoomTextDefaultValue = "59200 km";

        /// <summary>
        /// Sky Zoom Text Default Value
        /// </summary>
        public const string SkyZoomTextDefaultValue = "60:00:00";

        #endregion Perspective Properties

        #region Persistence

        /// <summary>
        /// XmlNamespace for serialization/de-serialization of Workbook map
        /// </summary>
        public const string XmlNamespace = "Microsoft.Research.Wwt.Excel.Addin.WorkbooMap";

        /// <summary>
        /// XmlRootName for serialization/de-serialization of Workbook map
        /// </summary>
        public const string XmlRootName = "WorkbookMap";

        /// <summary>
        /// XmlNamespace for serialization/de-serialization of Group
        /// </summary>
        public const string GroupXmlNamespace = "Microsoft.Research.Wwt.Excel.Common";

        /// <summary>
        /// XmlRootName for serialization/de-serialization of Group
        /// </summary>
        public const string GroupXmlRootName = "Group";

        /// <summary>
        /// XmlNamespace for serialization/de-serialization of Viewpoint map
        /// </summary>
        public const string ViewpointMapXmlNamespace = "Microsoft.Research.Wwt.Excel.Addin.ViewpointMap";

        /// <summary>
        /// XmlRootName for serialization/de-serialization of ViewpointMap
        /// </summary>
        public const string ViewpointMapRootName = "ViewpointMap";

        #endregion Persistence

        /// <summary>
        /// Display Name property of the product in registry.
        /// </summary>
        public const string DisplayNameProperty = "DisplayName";

        /// <summary>
        /// Display Version property of the product in registry.
        /// </summary>
        public const string DisplayVersionProperty = "DisplayVersion";

        /// <summary>
        /// Invalid characters in Name.
        /// </summary>
        public const string InvalidNameCharactersPattern = "[ !@#$%^&*()\\-+={}:;'/><,~`\"|]";

        /// <summary>
        /// Regular expression to identify entire row selection as part of selection range
        /// </summary>
        public const string EntireRowPattern = @"^\$\d+:\$\d+$";

        /// <summary>
        /// Regular expression to identify entire column selection as part of selection range
        /// </summary>
        public const string EntireColumnPattern = @"^\$[A-Z]+:\$[A-Z]+$";

        /// <summary>
        /// Regular expression to identify Hex Character Pattern in LCAPI response 
        /// </summary>
        public const string HexCharacterPattern = @"\p{C}+";

        /// <summary>
        /// Starts with number or a dot (.) regular expression pattern.
        /// </summary>
        public const string StartsWithDigitOrDotPattern = "^[0-9.]";

        /// <summary>
        /// Default Layer name.
        /// </summary>
        public const string DefaultLayerName = "WWTLayer";

        /// <summary>
        /// Look at value for earth.
        /// </summary>
        public const string EarthLookAt = "Earth";

        /// <summary>
        /// Look at value for SolarSystem.
        /// </summary>
        public const string SolarSystemLookAt = "SolarSystem";

        /// <summary>
        /// Reference frame value for earth.
        /// </summary>
        public const string EarthReferenceFrame = "Earth";

        /// <summary>
        /// Look at value for sky.
        /// </summary>
        public const string SkyLookAt = "Sky";

        /// <summary>
        /// Reference frame value for sky.
        /// </summary>
        public const string SkyReferenceFrame = "Sky";

        /// <summary>
        /// Reference frame name of Sun.
        /// </summary>
        public const string SunFrameName = "Sun";

        /// <summary>
        /// Reference frame path for earth.
        /// </summary>
        public const string EarthFramePath = "/Sun/Earth";

        /// <summary>
        /// Reference frame path for Sun.
        /// </summary>
        public const string SunFramePath = "/Sun";

        /// <summary>
        /// Reference frame path for earth.
        /// </summary>
        public const string SkyFramePath = "/Sky";

        /// <summary>
        /// Call out animation
        /// </summary>
        public const string CallOutAnimation = "CallOutLoadAnimation";

        /// <summary>
        /// Show highlight animation
        /// </summary>
        public const string ShowHighlightAnimation = "ShowHighlightAnimation";

        /// <summary>
        /// Hide highlight animation
        /// </summary>
        public const string HideHighlightAnimation = "HideHighlightAnimation";

        /// <summary>
        /// Base version for WWT
        /// </summary>
        public const string WWTBaseVersion = "3.0.57";

        /// <summary>
        /// Call out time interval(in seconds) for visibility setting.
        /// </summary>
        public const double CalloutTimerInterval = 7;

        /// <summary>
        /// Call out time interval(in seconds) for visibility setting.
        /// </summary>
        public const double ShowHighlightCalloutTimerInterval = 3;

        /// <summary>
        /// Feedback Link
        /// </summary>
        public const string FeedbackLink = "http://go.microsoft.com/?linkid=9788731";

        /// <summary>
        /// Visit a forum link
        /// </summary>
        public const string VisitForumLink = "http://go.microsoft.com/?linkid=9788730";

        /// <summary>
        /// Contact us link
        /// </summary>
        public const string ContactUsLink = "mailto:wwtefbk@microsoft.com";

        /// <summary>
        /// Excel Add-in Help link
        /// </summary>
        public const string HelpLink = "http://go.microsoft.com/?linkid=9790002";

        /// <summary>
        /// Download link for updates
        /// </summary>
        public const string DownloadUpdatesLink = "http://go.microsoft.com/?linkid=9821416";

        /// <summary>
        /// Version header in the server response for downloading updates
        /// </summary>
        public const string DownloadUpdatesResponseVersionHeader = "ClientVersion";

        /// <summary>
        /// Url header in the server response for downloading updates
        /// </summary>
        public const string DownloadUpdatesResponseUrlHeader = "UpdateUrl";

        /// <summary>
        /// Settings Node name.
        /// </summary>
        public const string SettingNodeName = "Settings";

        /// <summary>
        /// Name of add element.
        /// </summary>
        public const string AddNodeName = "Add";

        /// <summary>
        /// Name of Key attribute.
        /// </summary>
        public const string KeyAttributeName = "Key";

        /// <summary>
        /// Name of value attribute.
        /// </summary>
        public const string ValueAttributeName = "Value";

        /// <summary>
        /// Gets the config file name.
        /// </summary>
        public const string ConfigFileName = "WWTExcelAdd-In.config";

        #region Layer Properties

        /// <summary>
        /// Default ColumnIndex for layer columns
        /// </summary>
        public const int DefaultColumnIndex = -1;

        /// <summary>
        /// Default ColumnIndex for name column
        /// </summary>
        public const int DefaultNameColumnIndex = 0;

        /// <summary>
        /// Default layer version
        /// </summary>
        public const int DefaultLayerVersion = 0;

        /// <summary>
        /// Default Time decay for layer
        /// </summary>
        public const double DefaultTimeDecay = 16;

        /// <summary>
        /// Default Scale factor for layer
        /// </summary>
        public const double DefaultScaleFactor = 8;

        /// <summary>
        /// Default Opacity for layer
        /// </summary>
        public const int DefaultOpacity = 1;

        /// <summary>
        /// Default marker index.
        /// </summary>
        public const int DefaultMarkerIndex = 0;

        /// <summary>
        /// Default color for layer
        /// </summary>
        public const string DefaultColor = "ARGBColor:255:255:255:255";

        /// <summary>
        /// Text which identifies LayerApi.
        /// </summary>
        public const string LCAPIElementName = "LayerApi";

        /// <summary>
        /// Text which identifies error in processing.
        /// </summary>
        public const string LCAPIErrorText = "Error";

        /// <summary>
        /// Text which Identifies error in connection
        /// </summary>
        public const string LCAPIConnectionErrorText = "IP Not Authorized";

        /// <summary>
        /// Local host name.
        /// </summary>
        public const string Localhost = "localhost";

        /// <summary>
        /// Color prefix for the color to string conversion
        /// </summary>
        public const string ColorPrefix = "ARGBColor:";

        /// <summary>
        /// Color separator for the ARGB color
        /// </summary>
        public const string ColorSeparator = ":";

        /// <summary>
        /// Command for create layer.
        /// </summary>
        public const string CreateLayerCommand = "http://{0}:5050/layerApi.aspx?cmd=new&name={1}&frame={2}";

        /// <summary>
        /// Command for update layer properties.
        /// </summary>
        public const string UpdateLayerCommand = "http://{0}:5050/layerApi.aspx?cmd=setprops&id={1}";

        /// <summary>
        /// Default xml header tag.
        /// </summary>
        public const string XmlHeaderTag = "<?xml version=\"1.0\" encoding=\"utf-8\"?>";

        /// <summary>
        /// Upload data command.
        /// </summary>
        public const string UploadDataCommand = "http://{0}:5050/layerApi.aspx?cmd=update&id={1}&nopurge=true&hasheader=true";

        /// <summary>
        /// Upload data without header command.
        /// </summary>
        public const string UploadDataWithoutHeaderCommand = "http://{0}:5050/layerApi.aspx?cmd=update&id={1}&nopurge=true";

        /// <summary>
        /// Set property command.
        /// </summary>
        public const string SetPropertyCommand = "http://{0}:5050/layerApi.aspx?cmd=setprop&id={1}&propname={2}&propvalue={3}";

        /// <summary>
        /// Get all layers list command.
        /// </summary>
        public const string GetAllLayersCommand = "http://{0}:5050/layerApi.aspx?cmd=layerlist";

        /// <summary>
        /// Get layer details command.
        /// </summary>
        public const string GetLayerDetailsCommand = "http://{0}:5050/layerApi.aspx?cmd=getprops&id={1}";

        /// <summary>
        /// Purge data command.
        /// </summary>
        public const string PurgeDataCommand = "http://{0}:5050/layerApi.aspx?cmd=update&id={1}&purgeall=true&nopurge=true";

        /// <summary>
        /// Activate Layer command.
        /// </summary>
        public const string ActivateLayerCommand = "http://{0}:5050/layerApi.aspx?cmd=activate&id={1}";

        /// <summary>
        /// Show Layer Manager command.
        /// </summary>
        public const string ShowLayerManagerCommand = "http://{0}:5050/layerApi.aspx?cmd=showlayermanager";

        /// <summary>
        /// Delete Layer command.
        /// </summary>
        public const string DeleteLayerCommand = "http://{0}:5050/layerApi.aspx?cmd=delete&id={1}";

        /// <summary>
        /// Get header data command.
        /// </summary>
        public const string GetHeaderDataCommand = "http://{0}:5050/layerApi.aspx?cmd=get&id={1}";

        /// <summary>
        /// Get layer data command.
        /// </summary>
        public const string GetLayerDataCommand = "http://{0}:5050/layerApi.aspx?cmd=get&id={1}";

        /// <summary>
        /// Get properties command.
        /// </summary>
        public const string GetPropertiesCommand = "http://{0}:5050/layerApi.aspx?cmd=getprops&id={1}";

        /// <summary>
        /// Set perspective properties command
        /// </summary>
        public const string SetCameraViewCommand = "http://{0}:5050/layerApi.aspx?cmd=state&instant=false&flyto={1},{2},{3},{4},{5}&timerate={6}&datetime={7}";

        /// <summary>
        /// Set perspective properties command with reference frame and view token
        /// </summary>
        public const string SetCameraViewCommandWithReferenceFrame = "http://{0}:5050/layerApi.aspx?cmd=state&instant=false&flyto={1},{2},{3},{4},{5},{6},{7}&timerate={8}&datetime={9}";

        /// <summary>
        /// Get perspective properties command
        /// </summary>
        public const string GetCameraViewCommand = "http://{0}:5050/layerApi.aspx?cmd=state";

        /// <summary>
        /// Gets the state of the WWT on the given IP address 
        /// </summary>
        public const string GetWWTInstalledStateCommand = "http://{0}:5050/layerApi.aspx?cmd=version";

        /// <summary>
        /// Create layer group command.
        /// </summary>
        public const string CreateLayerGroupCommand = "http://{0}:5050/layerApi.aspx?cmd=group&frame={1}&name={2}";

        /// <summary>
        /// Set mode command 
        /// </summary>
        public const string SetModeCommand = "http://{0}:5050/layerApi.aspx?cmd=mode&lookat={1}";

        /// <summary>
        /// Notify layer command 
        /// </summary>
        public const string NotifyLayerCommand = "http://{0}:5050/layerApi.aspx?cmd=notify&notifytype=layer&notifytimeout=180000&notifyrate=100&id={1}&version={2}";

        /// <summary>
        /// Status attribute name
        /// </summary>
        public const string StatusAttribute = "Status";

        /// <summary>
        /// Version attribute name
        /// </summary>
        public const string VersionAttribute = "Version";

        /// <summary>
        /// Attribute name of ID.
        /// </summary>
        public const string IDAttribute = "ID";

        /// <summary>
        /// Attribute name of Name.
        /// </summary>
        public const string NameAttribute = "Name";

        /// <summary>
        /// Attribute name of Enabled.
        /// </summary>
        public const string EnabledAttribute = "Enabled";

        /// <summary>
        /// Attribute name of new layer.
        /// </summary>
        public const string NewLayerIDAttribute = "NewLayerID";

        /// <summary>
        /// Layer element name.
        /// </summary>
        public const string LayerElementNodeName = "Layer";

        /// <summary>
        /// Latitude column name.
        /// </summary>
        public const string LatColumnAttributeName = "LatColumn";

        /// <summary>
        /// Longitude column name.
        /// </summary>
        public const string LngColumnAttributeName = "LngColumn";

        /// <summary>
        /// Geometry Column Name
        /// </summary>
        public const string GeometryColumnAttributeName = "GeometryColumn";

        /// <summary>
        /// Color Map Column Name
        /// </summary>
        public const string ColorMapColumnAttributeName = "ColorMapColumn";

        /// <summary>
        /// Alt Column Name
        /// </summary>
        public const string AltColumnAttributeName = "AltColumn";

        /// <summary>
        /// StartDate Column Name
        /// </summary>
        public const string StartDateColumnAttributeName = "StartDateColumn";

        /// <summary>
        /// EndDate Column Name
        /// </summary>
        public const string EndDateColumnAttributeName = "EndDateColumn";

        /// <summary>
        /// Size Column Name
        /// </summary>
        public const string SizeColumnAttributeName = "SizeColumn";

        /// <summary>
        /// Name Column Name
        /// </summary>
        public const string NameColumnAttributeName = "NameColumn";

        /// <summary>
        /// Decay Name
        /// </summary>
        public const string DecayAttributeName = "Decay";

        /// <summary>
        /// Scale Factor Name
        /// </summary>
        public const string ScaleFactorAttributeName = "ScaleFactor";

        /// <summary>
        /// Opacity Name
        /// </summary>
        public const string OpacityAttributeName = "Opacity";

        /// <summary>
        /// StartTime Name
        /// </summary>
        public const string StartTimeAttributeName = "StartTime";

        /// <summary>
        /// EndTime Name
        /// </summary>
        public const string EndTimeAttributeName = "EndTime";

        /// <summary>
        /// FadeSpan Name
        /// </summary>
        public const string FadeSpanAttributeName = "FadeSpan";

        /// <summary>
        /// ColorValue Name
        /// </summary>
        public const string ColorValueAttributeName = "ColorValue";

        /// <summary>
        /// AltType Name
        /// </summary>
        public const string AltTypeAttributeName = "AltType";

        /// <summary>
        /// MarkerScale Name
        /// </summary>
        public const string MarkerScaleAttributeName = "MarkerScale";

        /// <summary>
        /// AltUnit Name
        /// </summary>
        public const string AltUnitAttributeName = "AltUnit";

        /// <summary>
        /// CartesianScale Name
        /// </summary>
        public const string CartesianScaleAttributeName = "CartesianScale";

        /// <summary>
        /// RA Unit Name
        /// </summary>
        public const string RAUnitAttributeName = "RaUnits";

        /// <summary>
        /// PointScaleType Name
        /// </summary>
        public const string PointScaleTypeAttributeName = "PointScaleType";

        /// <summary>
        /// FadeType Name
        /// </summary>
        public const string FadeTypeAttributeName = "FadeType";

        /// <summary>
        /// Marker type attribute name
        /// </summary>
        public const string MarkerTypeAttributeName = "PlotType";

        /// <summary>
        /// Marker index attribute name
        /// </summary>
        public const string MarkerIndexAttributeName = "MarkerIndex";

        /// <summary>
        /// Show Far Side attribute name
        /// </summary>
        public const string ShowFarSideAttributeName = "ShowFarSide";

        /// <summary>
        /// Co-ordinate type attribute name
        /// </summary>
        public const string CoordinateTypeAttributeName = "CoordinatesType";

        /// <summary>
        /// X axis column attribute name
        /// </summary>
        public const string XAxisColumnAttributeName = "XAxisColumn";

        /// <summary>
        /// Y axis column attribute name
        /// </summary>
        public const string YAxisColumnAttributeName = "YAxisColumn";

        /// <summary>
        /// Z axis column attribute name
        /// </summary>
        public const string ZAxisColumnAttributeName = "ZAxisColumn";

        /// <summary>
        /// Reverse X axis column attribute name
        /// </summary>
        public const string ReverseXAxisColumnAttributeName = "XAxisReverse";

        /// <summary>
        /// Reverse Y axis column attribute name
        /// </summary>
        public const string ReverseYAxisColumnAttributeName = "YAxisReverse";

        /// <summary>
        /// Reverse Z axis column attribute name
        /// </summary>
        public const string ReverseZAxisColumnAttributeName = "ZAxisReverse";

        /// <summary>
        /// TimeSeries Name
        /// </summary>
        public const string TimeSeriesAttributeName = "TimeSeries";

        /// <summary>
        /// This is the XPath for searching the first reference frame.
        /// </summary>
        public const string XPathStringForLayerList = "//LayerApi/LayerList";

        /// <summary>
        /// This is the XPath for view state.
        /// </summary>
        public const string XPathStringForViewState = "//ViewState";

        /// <summary>
        /// Look at Attribute
        /// </summary>
        public const string LookatAttribute = "lookat";

        /// <summary>
        /// Lat Attribute
        /// </summary>
        public const string LatAttribute = "lat";

        /// <summary>
        /// Long Attribute
        /// </summary>
        public const string LongAttribute = "lng";

        /// <summary>
        /// Zoom Attribute
        /// </summary>
        public const string ZoomAttribute = "zoom";

        /// <summary>
        /// Rotation Attribute
        /// </summary>
        public const string RotationAttribute = "rotation";

        /// <summary>
        /// Angle Attribute
        /// </summary>
        public const string AngleAttribute = "angle";

        /// <summary>
        /// Reference Frame Node Name
        /// </summary>
        public const string ReferenceFrameElementName = "ReferenceFrame";

        /// <summary>
        /// Layer Group Node Name
        /// </summary>
        public const string LayerGroupElementName = "LayerGroup";

        /// <summary>
        /// Sun reference path.
        /// </summary>
        public const string SunReferencePath = "/Sun";

        /// <summary>
        /// TimeRate Attribute
        /// </summary>
        public const string TimeRateAttribute = "timerate";

        /// <summary>
        /// Observing Time Attribute
        /// </summary>
        public const string ObservingTimeAttribute = "time";

        /// <summary>
        /// Ra Attribute
        /// </summary>
        public const string RightAscentionAttribute = "ra";

        /// <summary>
        /// Dec Attribute
        /// </summary>
        public const string DeclinationAttribute = "dec";

        /// <summary>
        /// Reference Frame Attribute
        /// </summary>
        public const string ReferenceFrameAttribute = "ReferenceFrame";

        /// <summary>
        /// ZoomText Attribute
        /// </summary>
        public const string ZoomTextAttribute = "ZoomText";

        /// <summary>
        /// ViewToken Attribute
        /// </summary>
        public const string ViewTokenAttribute = "ViewToken";

        /// <summary>
        /// Gets default start time for layer
        /// </summary>
        public static DateTime DefaultStartTime
        {
            get
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>
        /// Gets default end time for layer
        /// </summary>
        public static DateTime DefaultEndTime
        {
            get
            {
                return DateTime.MaxValue;
            }
        }

        /// <summary>
        /// Gets default fade span for layer
        /// </summary>
        public static TimeSpan DefaultFadeSpan
        {
            get
            {
                return new TimeSpan(00, 00, 00);
            }
        }

        /// <summary>
        /// Gets the default response xml string.
        /// </summary>
        public static string DefaultErrorResponse
        {
            get
            {
                return (new System.Xml.Linq.XElement(
                    Constants.LCAPIElementName,
                    new System.Xml.Linq.XElement(Constants.StatusAttribute, Constants.LCAPIErrorText)))
                    .ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
            }
        }

        #endregion Layer Properties

        public static List<string> LatSearchList
        {
            get
            {
                return new List<string>() { "lat", "lt", "latitude" };
            }
        }

        public static List<string> LonSearchList
        {
            get
            {
                return new List<string>() { "lon", "ln", "longitude", "lng" };
            }
        }
    }
}
