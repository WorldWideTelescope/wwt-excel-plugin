//-----------------------------------------------------------------------
// <copyright file="WWTManager.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Research.Wwt.Excel.Common.Properties;
using Microsoft.Win32;

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Business layer class which is responsible for handling all the communication with WWT.
    /// </summary>
    public class WWTManager
    {
        /// <summary>
        /// IWWTRequest which sends request to WWT.
        /// </summary>
        private static IWWTRequest request;

        /// <summary>
        /// Initializes a new instance of the WWTManager class. It takes IWWTRequest as parameter and initializes request member.
        /// </summary>
        /// <param name="wwtRequest">IWWTRequest object</param>
        public WWTManager(IWWTRequest wwtRequest)
        {
            request = wwtRequest;
        }

        #region Public Static Methods
        /// <summary>
        /// This function is used to create a layer in WWT using LCAPI.
        /// </summary>
        /// <param name="layerName">
        /// Name of the layer.
        /// </param>
        /// <param name="frame">
        /// Name of the frame under which the layer has to be created.
        /// </param>
        /// <param name="headerData">
        /// Header data in comma separated string.
        /// </param>
        /// <returns>
        /// Return the layer ID of the layer which got created in WWT.
        /// </returns>
        public static string CreateLayer(string layerName, string frame, string headerData)
        {
            string layerId = string.Empty;
            if (!string.IsNullOrEmpty(layerName))
            {
                // Create a new layer with a data format
                string url = string.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                    Constants.CreateLayerCommand,
                    Globals.TargetMachine.MachineIP,
                    Encode(layerName),
                    Encode(frame));

                string response = request.Send(url, headerData, false);
                layerId = ParseLayerID(response);
            }

            return layerId;
        }

        /// <summary>
        /// This function is used to update the layer header in WWT using LCAPI.
        /// </summary>
        /// <param name="layerDetails">
        /// Details of the layer.
        /// </param>
        /// <param name="isConsumeException">
        /// Whether to consume exception?
        /// </param>
        /// <param name="isTimeSeriesRequired">
        /// Whether to time series is required or not.
        /// Time series is set explicitly only when a layer is created.
        /// </param>
        /// <returns>
        /// True if the layer is updated; otherwise false.
        /// </returns>
        public static bool UpdateLayer(Layer layerDetails, bool isConsumeException, bool isTimeSeriesRequired)
        {
            bool isValid = false;
            if (layerDetails != null)
            {
                // Create a new layer with a data format
                string url = string.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                    Constants.UpdateLayerCommand,
                    Globals.TargetMachine.MachineIP,
                    layerDetails.ID);

                string header = Constants.XmlHeaderTag + GetLayerProperties(layerDetails, isTimeSeriesRequired);
                string response = request.Send(url, header, isConsumeException);
                isValid = IsValidResponse(response);
            }

            return isValid;
        }

        /// <summary>
        /// This function is used to upload the data to WWT. 
        /// Before uploading the data the existing layer data is purged and then new data is uploaded.
        /// </summary>
        /// <param name="layerId">
        /// ID of the layer.
        /// </param>
        /// <param name="data">
        /// Data which has to be uploaded to WWT.
        /// </param>
        /// <param name="isConsumeException">
        /// Whether to consume exception?
        /// </param>
        /// <returns>
        /// True if the layer has updated; Otherwise false.
        /// </returns>
        public static bool UploadDataInWWT(string layerId, string[] data, bool isConsumeException)
        {
            bool hasUpdated = true;
            if (!string.IsNullOrEmpty(layerId) && data != null)
            {
                // Purge existing data from the layer in WWT.
                hasUpdated = PurgeExistingData(layerId);
                if (hasUpdated)
                {
                    string url = string.Format(
                        System.Globalization.CultureInfo.InvariantCulture,
                        Constants.UploadDataCommand,
                        Globals.TargetMachine.MachineIP,
                        layerId);

                    int count = 0;
                    foreach (string item in data)
                    {
                        string response = request.Send(url, item, isConsumeException);
                        if (!IsValidResponse(response) && hasUpdated)
                        {
                            hasUpdated = false;
                            break;
                        }

                        count++;

                        // Set the Uri to upload command without header from second data onwards
                        if (count == 1)
                        {
                            url = string.Format(
                            System.Globalization.CultureInfo.InvariantCulture,
                            Constants.UploadDataWithoutHeaderCommand,
                            Globals.TargetMachine.MachineIP,
                            layerId);
                        }
                    }
                }
            }

            return hasUpdated;
        }

        /// <summary>
        /// Set the perspective properties in WWT
        /// </summary>
        /// <param name="perspective">Perspective object</param>
        /// <returns>True if the update request is sent successfully</returns>
        public static void SetCameraView(Perspective perspective)
        {
            if (perspective != null)
            {
                string url = string.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                    Constants.SetModeCommand,
                    Globals.TargetMachine.MachineIP,
                    perspective.LookAt);

                if (IsValidResponse(request.Send(url, string.Empty, false)))
                {
                    if (perspective.LookAt.Equals(Constants.EarthLookAt, StringComparison.OrdinalIgnoreCase) ||
                        perspective.LookAt.Equals(Constants.SkyLookAt, StringComparison.OrdinalIgnoreCase))
                    {
                        url = string.Format(
                            System.Globalization.CultureInfo.InvariantCulture,
                            Constants.SetCameraViewCommand,
                            Globals.TargetMachine.MachineIP,
                            perspective.HasRADec ? perspective.Declination : perspective.Latitude,
                            perspective.HasRADec ? perspective.RightAscention : perspective.Longitude,
                            perspective.Zoom,
                            perspective.Rotation,
                            perspective.LookAngle,
                            perspective.TimeRate,
                            perspective.ObservingTime);
                    }
                    else
                    {
                        url = string.Format(
                            System.Globalization.CultureInfo.InvariantCulture,
                            Constants.SetCameraViewCommandWithReferenceFrame,
                            Globals.TargetMachine.MachineIP,
                            perspective.HasRADec ? perspective.Declination : perspective.Latitude,
                            perspective.HasRADec ? perspective.RightAscention : perspective.Longitude,
                            perspective.Zoom,
                            perspective.Rotation,
                            perspective.LookAngle,
                            perspective.ReferenceFrame,
                            perspective.ViewToken,
                            perspective.TimeRate,
                            perspective.ObservingTime);
                    }

                    request.Send(url, string.Empty, false);
                }
            }
        }

        /// <summary>
        /// Set the mode in WWT
        /// </summary>
        /// <param name="lookAt">look At value</param>
        public static void SetMode(string lookAt)
        {
            string url = string.Format(
                System.Globalization.CultureInfo.InvariantCulture,
                Constants.SetModeCommand,
                Globals.TargetMachine.MachineIP,
                lookAt);

            request.Send(url, string.Empty, false);
        }

        /// <summary>
        /// Retrieve the perspective properties from WWT
        /// </summary>
        /// <returns>Perspective object containing camera details</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Justification = "LCAPI method to get the perspective details")]
        public static Perspective GetCameraView()
        {
            Perspective perspective = null;

            string url = string.Format(
                System.Globalization.CultureInfo.InvariantCulture,
                Constants.GetCameraViewCommand,
                Globals.TargetMachine.MachineIP);

            string response = request.Send(url, string.Empty, false);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response);
            XmlNode viewStateNode = doc.SelectSingleNode(Constants.XPathStringForViewState);

            if (viewStateNode != null)
            {
                if (GetAttributeValue(viewStateNode, Constants.LookatAttribute).Equals(Constants.SkyLookAt, StringComparison.OrdinalIgnoreCase))
                {
                    perspective = new Perspective(Common.Constants.SkyLookAt, Common.Constants.SkyReferenceFrame, true, Common.Constants.LatitudeDefaultValue, Common.Constants.LongitudeDefaultValue, Common.Constants.ZoomDefaultValue, Common.Constants.RotationDefaultValue, Common.Constants.LookAngleDefaultValue, DateTime.Now.ToString(), Common.Constants.TimeRateDefaultValue, Common.Constants.SkyZoomTextDefaultValue, string.Empty);
                    perspective.RightAscention = GetAttributeValue(viewStateNode, Constants.RightAscentionAttribute);
                    perspective.Declination = GetAttributeValue(viewStateNode, Constants.DeclinationAttribute);
                }
                else
                {
                    perspective = new Perspective(Common.Constants.EarthLookAt, Common.Constants.EarthLookAt, false, Common.Constants.LatitudeDefaultValue, Common.Constants.LongitudeDefaultValue, Common.Constants.ZoomDefaultValue, Common.Constants.RotationDefaultValue, Common.Constants.LookAngleDefaultValue, DateTime.Now.ToString(), Common.Constants.TimeRateDefaultValue, Common.Constants.EarthZoomTextDefaultValue, string.Empty);
                    perspective.Latitude = GetAttributeValue(viewStateNode, Constants.LatAttribute);
                    perspective.Longitude = GetAttributeValue(viewStateNode, Constants.LongAttribute);
                    perspective.LookAngle = GetAttributeValue(viewStateNode, Constants.AngleAttribute);
                }

                perspective.LookAt = GetAttributeValue(viewStateNode, Constants.LookatAttribute);
                perspective.ReferenceFrame = GetAttributeValue(viewStateNode, Constants.ReferenceFrameAttribute);
                perspective.ViewToken = GetAttributeValue(viewStateNode, Constants.ViewTokenAttribute);
                perspective.ZoomText = GetAttributeValue(viewStateNode, Constants.ZoomTextAttribute);
                perspective.Rotation = GetAttributeValue(viewStateNode, Constants.RotationAttribute);
                perspective.Zoom = GetAttributeValue(viewStateNode, Constants.ZoomAttribute);
                perspective.TimeRate = GetAttributeValue(viewStateNode, Constants.TimeRateAttribute);
                perspective.ObservingTime = GetAttributeValue(viewStateNode, Constants.ObservingTimeAttribute);
            }

            return perspective;
        }

        /// <summary>
        /// This function is used to set the property value in WWT.
        /// </summary>
        /// <param name="layerId">
        /// ID of the layer.
        /// </param>
        /// <param name="propname">
        /// Property name.
        /// </param>
        /// <param name="propvalue">
        /// Property Value.
        /// </param>
        /// <exception cref="CustomException">
        /// If WWT is not open, then we throw CustomException with message WWT not Open. 
        /// </exception>
        public static void SetProperty(string layerId, string propname, string propvalue)
        {
            string url = string.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                Constants.SetPropertyCommand,
                Globals.TargetMachine.MachineIP,
                layerId,
                propname,
                propvalue);
            request.Send(url, string.Empty, true);
        }

        /// <summary>
        /// Activates WWT layer
        /// </summary>
        /// <param name="layerId">layer ID of the layer to be activated</param>
        public static void ActivateLayer(string layerId)
        {
            string url = string.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                Constants.ActivateLayerCommand,
                Globals.TargetMachine.MachineIP,
                layerId);

            request.Send(url, string.Empty, true);
        }

        /// <summary>
        /// Shows the layer manager pane in WorldWide Telescope.
        /// </summary>
        public static void ShowLayerManager()
        {
            string url = string.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                Constants.ShowLayerManagerCommand,
                Globals.TargetMachine.MachineIP);

            request.Send(url, string.Empty, true);
        }

        /// <summary>
        /// Deletes the layer from WWT
        /// </summary>
        /// <param name="layerId">Selected layer id</param>
        public static void DeleteLayer(string layerId)
        {
            if (!string.IsNullOrEmpty(layerId))
            {
                string url = string.Format(
                             System.Globalization.CultureInfo.InvariantCulture,
                             Constants.DeleteLayerCommand,
                             Globals.TargetMachine.MachineIP,
                             layerId);

                request.Send(url, string.Empty, false);
            }
        }

        /// <summary>
        /// Sends a request for notifications from WWT in case of any change to the given layer properties.
        /// </summary>
        /// <param name="layerId">Id of the layer</param>
        /// <param name="layerVersion">Current version of the layer</param>
        /// <returns>
        /// Updated Version number. In case of no change, same version will be returned. In case of any change, latest version will be returned.
        /// In case of any error, i.e. because WWT is closed or layer is deleted, -1 will be returned as version.
        /// </returns>
        public static int CreateLayerNotification(string layerId, int layerVersion, CancellationToken cancellationToken)
        {
            if (!string.IsNullOrEmpty(layerId))
            {
                string url = string.Format(System.Globalization.CultureInfo.InvariantCulture, Constants.NotifyLayerCommand, Globals.TargetMachine.MachineIP, layerId, layerVersion);
                Task<string> webTask = SendAsync(url, cancellationToken);
                string response = webTask.Result;

                if (IsValidResponse(response))
                {
                    layerVersion = ParseLayerVersion(response);
                }
                else if (!IsTimeoutResponse(response))
                {
                    // Set the version as -1 which is indicate that the layer is not available in WWT.
                    layerVersion = -1;
                }
            }
            else
            {
                // If IDis not there, then notification cannot happen and should not continue.
                layerVersion = -1;
            }

            return layerVersion;
        }

        /// <summary>
        /// Get all WWT layers ID's for the specified frame.
        /// </summary>
        /// <param name="frame">
        /// Name of the frame for which we have to retrieve the layers.
        /// </param>
        /// <param name="isConsumeException">
        /// Whether to consume exception?
        /// </param>
        /// <returns>
        /// List of all layers which belong to the frame in focus.
        /// </returns>
        /// <exception cref="CustomException">
        /// If WWT is not open, then we throw CustomException with message WWT not Open. 
        /// </exception>
        public static ICollection<string> GetAllWWTLayerIds(string frame, bool isConsumeException)
        {
            List<string> layers = new List<string>();

            string url = string.Format(
                System.Globalization.CultureInfo.InvariantCulture,
                Constants.GetAllLayersCommand,
                Globals.TargetMachine.MachineIP);

            string response = request.Send(url, string.Empty, isConsumeException);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response);

            foreach (XmlElement element in doc.SelectNodes(Constants.XPathStringForLayerList))
            {
                if (string.CompareOrdinal(element.Attributes[Constants.NameAttribute].Value, frame) == 0)
                {
                    foreach (XmlElement child in element.ChildNodes)
                    {
                        if (child.Attributes[Constants.IDAttribute] != null)
                        {
                            layers.Add(child.Attributes[Constants.IDAttribute].Value);
                        }
                    }
                }
            }

            return layers;
        }

        /// <summary>
        /// This function retrieves the layers from WWT in the format for parent child relationship.
        /// </summary>
        /// <param name="isConsumeException">
        /// Whether to consume exception?
        /// </param>
        /// <returns>
        /// List of all layers
        /// </returns>
        public static ICollection<Group> GetAllWWTGroups(bool isConsumeException)
        {
            List<Group> groups = new List<Group>();

            string url = string.Format(
                System.Globalization.CultureInfo.InvariantCulture,
                Constants.GetAllLayersCommand,
                Globals.TargetMachine.MachineIP);

            string response = request.Send(url, string.Empty, isConsumeException);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response);

            XmlNode node = doc.SelectSingleNode(Constants.XPathStringForLayerList);
            if (node != null && node.HasChildNodes)
            {
                foreach (XmlElement element in node.ChildNodes)
                {
                    if (element.Name == Constants.ReferenceFrameElementName || element.Name == Constants.LayerGroupElementName)
                    {
                        Group childGroup = GetGroup(element, null);
                        if (childGroup != null)
                        {
                            groups.Add(childGroup);
                        }
                    }
                }
            }

            return groups;
        }

        /// <summary>
        /// This function is used to create the layer group under the specified parent.
        /// </summary>
        /// <param name="name">
        /// Name of the layer group.
        /// </param>
        /// <param name="parentName">
        /// Name of the parent group.
        /// </param>
        public static void CreateLayerGroup(string name, string parentName)
        {
            if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(parentName))
            {
                // Create a new layer with a data format
                string url = string.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                    Constants.CreateLayerGroupCommand,
                    Globals.TargetMachine.MachineIP,
                    Encode(parentName),
                    Encode(name));

                request.Send(url, string.Empty, false);
            }
        }

        /// <summary>
        /// This function is used to load the layer details from WWT.
        /// </summary>
        /// <param name="layerId">
        /// ID of the layer.
        /// </param>
        /// <param name="group">
        /// Group of the layer.
        /// </param>
        /// <param name="isConsumeException">
        /// Whether to consume exception?
        /// </param>
        /// <returns>
        /// Layer details.
        /// </returns>
        public static Layer GetLayerDetails(string layerId, Group group, bool isConsumeException)
        {
            string url = string.Format(
                System.Globalization.CultureInfo.InvariantCulture,
                Constants.GetLayerDetailsCommand,
                Globals.TargetMachine.MachineIP,
                layerId);
            string response = request.Send(url, string.Empty, isConsumeException);
            if (IsValidResponse(response))
            {
                return GetLayer(response, layerId, group);
            }

            return null;
        }

        /// <summary>
        /// Get header data for the layer
        /// </summary>
        /// <param name="layerId">Id of the layer</param>
        /// <returns>header data as list</returns>
        public static Collection<string> GetLayerHeader(string layerId)
        {
            Collection<string> headerData = new Collection<string>();

            string url = string.Format(
                System.Globalization.CultureInfo.InvariantCulture,
                Constants.GetHeaderDataCommand,
                Globals.TargetMachine.MachineIP,
                layerId);

            Stream stream = null;
            StreamReader streamReader = null;
            try
            {
                using (var client = new WebClient())
                {
                    stream = client.OpenRead(url);
                    streamReader = new StreamReader(stream);
                    headerData = new Collection<string>(SplitString(streamReader.ReadLine()));
                }
            }
            catch (WebException ex)
            {
                Logger.LogException(ex);
            }
            catch (IOException ex)
            {
                Logger.LogException(ex);
            }
            finally
            {
                if (stream != null)
                {
                    stream.Dispose();
                }
            }

            return headerData;
        }

        /// <summary>
        /// Gets layer data for the given layer id
        /// </summary>
        /// <param name="layerId">Layer id(Local in WWT/WWT)</param>
        /// <param name="consumeException">Boolean indicating whether to throw the exception</param>
        /// <returns>Object array of the data from WWT for selected layer</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1814:PreferJaggedArraysOverMultidimensional", MessageId = "Return", Justification = "We cannot use jagged array in this scenario because the excel Object model is designed to convert the value as [,].")]
        public static object[,] GetLayerData(string layerId, bool consumeException)
        {
            object[,] layerData = null;
            if (!string.IsNullOrWhiteSpace(layerId))
            {
                string url = string.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                    Constants.GetLayerDataCommand,
                    Globals.TargetMachine.MachineIP,
                    layerId);

                Stream stream = null;
                StreamReader streamReader = null;
                try
                {
                    using (var client = new WebClient())
                    {
                        stream = client.OpenRead(url);
                        streamReader = new StreamReader(stream);
                        if (streamReader != null)
                        {
                            int cols = 0;
                            Collection<string[]> data = new Collection<string[]>();
                            while (!streamReader.EndOfStream)
                            {
                                List<string> line = SplitString(streamReader.ReadLine());
                                cols = Math.Max(cols, line.Count);
                                data.Add(line.ToArray());
                            }
                            layerData = GetAllLayerData(data, cols);
                        }
                    }
                }
                catch (ArgumentException ex)
                {
                    Logger.LogException(ex);

                    // Consume as exception while reading data
                    if (!consumeException)
                    {
                        throw new CustomException(Properties.Resources.DefaultErrorMessage, true);
                    }
                }
                catch (WebException ex)
                {
                    Logger.LogException(ex);

                    // throw custom exception if unable to connect
                    if (!consumeException)
                    {
                        throw new CustomException(Properties.Resources.WWTNotOpenFailure, ex, true, ErrorCodes.Code100001);
                    }
                }
                catch (IOException ex)
                {
                    Logger.LogException(ex);

                    // Consume as exception while reading data
                    if (!consumeException)
                    {
                        throw new CustomException(Properties.Resources.DefaultErrorMessage, true);
                    }
                }
                catch (OutOfMemoryException ex)
                {
                    Logger.LogException(ex);

                    // Consume as exception while reading data
                    if (!consumeException)
                    {
                        throw new CustomException(Properties.Resources.DefaultErrorMessage, true);
                    }
                }
                catch (ObjectDisposedException ex)
                {
                    Logger.LogException(ex);

                    // Consume as exception while reading data
                    if (!consumeException)
                    {
                        throw new CustomException(Properties.Resources.DefaultErrorMessage, true);
                    }
                }
                finally
                {
                    if (stream != null)
                    {
                        stream.Dispose();
                    }
                }
            }
            return layerData;
        }

        /// <summary>
        /// This function is used to check if the given layer id is present in WWT or not.
        /// </summary>
        /// <param name="layerId">
        /// ID of the layer in focus.
        /// </param>
        /// <returns>
        /// True, if the layer is present in WWT;Otherwise false.
        /// </returns>
        public static bool IsValidLayer(string layerId)
        {
            bool isValid = false;

            if (!string.IsNullOrEmpty(layerId))
            {
                ICollection<Group> groups = GetAllWWTGroups(true);
                isValid = IsValidLayer(layerId, groups);
            }

            return isValid;
        }

        /// <summary>
        /// This function is used to check if the given layer id is present in WWT or not.
        /// </summary>
        /// <param name="layerId">
        /// ID of the layer in focus.
        /// </param>
        /// <param name="groups">
        /// List of WWT Groups.
        /// </param>
        /// <returns>
        /// True, if the layer is present in WWT;Otherwise false.
        /// </returns>
        public static bool IsValidLayer(string layerId, ICollection<Group> groups)
        {
            bool isValid = false;
            if (!string.IsNullOrEmpty(layerId) && groups != null)
            {
                foreach (Group group in groups)
                {
                    if (IsValidLayer(layerId, group))
                    {
                        isValid = true;
                        break;
                    }
                }
            }

            return isValid;
        }

        /// <summary>
        /// This function is used to check if the given group is present in WWT or not.
        /// </summary>
        /// <param name="group">
        /// Instance of group.
        /// </param>
        /// <param name="wwtGroups">
        /// List of WWT Groups.
        /// </param>
        /// <returns>
        /// True, if the group is present in WWT;Otherwise false.
        /// </returns>
        public static bool IsValidGroup(Group group, ICollection<Group> wwtGroups)
        {
            bool isValid = false;
            if (group != null && wwtGroups != null)
            {
                foreach (Group wwtGroup in wwtGroups)
                {
                    if (IsValidGroup(group, wwtGroup))
                    {
                        isValid = true;
                        break;
                    }
                }
            }

            return isValid;
        }

        /// <summary>
        /// Checks if the machine IP is valid
        /// </summary>
        /// <param name="machineIP">Target machine IP</param>
        /// <param name="cosumeException">Flag for consuming exception</param>
        /// <returns>True if the machine has valid IP</returns>
        public static bool IsValidMachine(string machineIP, bool cosumeException)
        {
            bool isValid = false;
            if (!string.IsNullOrEmpty(machineIP))
            {
                string url = string.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                    Constants.GetWWTInstalledStateCommand,
                    machineIP);
                string response = request.Send(url, string.Empty, cosumeException);

                isValid = IsValidVersion(response);

                // If the machine is valid, but WWT is not running or has an older version,it throws 
                // a custom exception
                if (!isValid && !cosumeException)
                {
                    throw new CustomException(Properties.Resources.WWTOlderVersionError, true, ErrorCodes.Code100003);
                }
            }
            return isValid;
        }

        /// <summary>
        /// Check the server for an updated version of WWTE Excel Addin
        /// </summary>
        /// <returns>True if an updated version is available on the server</returns>
        public static string CheckForUpdates()
        {
            string updateUrl = string.Empty;

            try
            {
                using (WebClient webClient = new WebClient())
                {
                    string data = webClient.DownloadString(Constants.DownloadUpdatesLink);
                    string[] lines = data.Split(new char[] { '\n' });
                    if (!lines[0].StartsWith(Common.Constants.DownloadUpdatesResponseVersionHeader, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new CustomException(Resources.DownloadUpdatesResponseError);
                    }
                    if (!lines[1].StartsWith(Common.Constants.DownloadUpdatesResponseUrlHeader, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new CustomException(Resources.DownloadUpdatesResponseError);
                    }

                    Version currentVersion = Version.Parse(lines[0].Substring(lines[0].IndexOf(':') + 1).Trim());
                    string version = GetAddinInstalledVersion();
                    Version installedVersion = new Version();
                    if (!string.IsNullOrEmpty(version))
                    {
                        installedVersion = Version.Parse(version);
                    }
                    if (currentVersion > installedVersion)
                    {
                        updateUrl = lines[1].Substring(lines[1].IndexOf(':') + 1).Trim();
                    }
                }
            }
            catch (WebException ex)
            {
                Logger.LogException(ex);
            }
            catch (System.Security.SecurityException ex)
            {
                Logger.LogException(ex);
            }
            catch (System.IO.IOException ex)
            {
                Logger.LogException(ex);
            }
            catch (System.UnauthorizedAccessException ex)
            {
                Logger.LogException(ex);
            }
            catch (ArgumentException ex)
            {
                Logger.LogException(ex);
            }
            catch (NotSupportedException ex)
            {
                Logger.LogException(ex);
            }
            return updateUrl;
        }

        /// <summary>
        /// Downloads MSI file to get the updated WWT version and it to 
        /// temporary folder.
        /// </summary>
        /// <param name="uri">URI from where the file has to be downloaded</param>
        /// <param name="filename">Download file name with path</param>
        /// <returns>Returns if the file is downloaded or not.</returns>
        public static bool DownloadFile(Uri uri, string filename)
        {
            try
            {
                if (uri != null & !string.IsNullOrWhiteSpace(filename))
                {
                    if (File.Exists(filename))
                    {
                        File.Delete(filename);
                    }

                    if (uri.IsFile)
                    {
                        string source = uri.GetComponents(UriComponents.Path, UriFormat.SafeUnescaped);
                        File.Copy(source, filename);
                        return true;
                    }
                    using (WebClient client = new WebClient())
                    {
                        client.DownloadFile(uri, filename);
                    }
                }
            }
            catch (WebException ex)
            {
                Logger.LogException(ex);
            }
            catch (System.IO.IOException ex)
            {
                Logger.LogException(ex);
            }
            catch (System.UnauthorizedAccessException ex)
            {
                Logger.LogException(ex);
            }
            catch (ArgumentException ex)
            {
                Logger.LogException(ex);
            }
            catch (NotSupportedException ex)
            {
                Logger.LogException(ex);
            }
            catch (InvalidOperationException ex)
            {
                Logger.LogException(ex);
            }
            return true;
        }
        #endregion

        #region Private static Methods
        /// <summary>
        /// This function is used to purge the existing data in the layer.
        /// </summary>
        /// <param name="layerId">
        /// ID of the layer.
        /// </param>
        private static bool PurgeExistingData(string layerId)
        {
            string url = string.Format(
                System.Globalization.CultureInfo.InvariantCulture,
                Constants.PurgeDataCommand,
                Globals.TargetMachine.MachineIP,
                layerId);

            return IsValidResponse(request.Send(url, string.Empty, true));
        }

        /// <summary>
        /// This function is used to Encode the input string.
        /// </summary>
        /// <param name="input">
        /// String which has to be encoded.
        /// </param>
        /// <returns>
        /// Encoded string.
        /// </returns>
        private static string Encode(string input)
        {
            return HttpUtility.UrlEncode(input);
        }

        /// <summary>
        /// This function is used to retrieve the layer id from the response.
        /// </summary>
        /// <param name="response">
        /// Response for the current operation.
        /// </param>
        /// <returns>
        /// New Layer ID.
        /// </returns>
        private static string ParseLayerID(string response)
        {
            string layerId = string.Empty;
            try
            {
                XElement root = XElement.Parse(response);
                layerId = root.Element(Constants.NewLayerIDAttribute).Value;
            }
            catch (XmlException ex)
            {
                Logger.LogException(ex);

                // Ignore error.
            }

            return layerId;
        }

        /// <summary>
        /// This function is used to retrieve the layer version from the response.
        /// </summary>
        /// <param name="response">Response for the current operation</param>
        /// <returns>Layer version</returns>
        private static int ParseLayerVersion(string response)
        {
            int layerVersion = 0;
            try
            {
                XElement root = XElement.Parse(response);

                if (root.Element(Constants.LayerElementNodeName) != null && root.Element(Constants.LayerElementNodeName).Attribute(Constants.VersionAttribute) != null)
                {
                    layerVersion = Convert.ToInt32(root.Element(Constants.LayerElementNodeName).Attribute(Constants.VersionAttribute).Value, CultureInfo.CurrentCulture);
                }
            }
            catch (XmlException ex)
            {
                Logger.LogException(ex);

                // Ignore error.
            }

            return layerVersion;
        }

        /// <summary>
        /// This function is used to create the layer properties from the response XML.
        /// </summary>
        /// <param name="response">
        /// Response xml.
        /// </param>
        /// <param name="layerId">
        /// Id of the layer.
        /// </param>
        /// <param name="group">
        /// Group of the layer.
        /// </param>
        /// <returns>
        /// Instance of layer class.
        /// </returns>
        private static Layer GetLayer(string response, string layerId, Group group)
        {
            Layer layer = new Layer();
            layer.ID = layerId;
            layer.Group = group;
            XElement root = XElement.Parse(response, LoadOptions.PreserveWhitespace);

            // Get All Attributes list  of Layers.
            var listOfAttributes = root.Element(Constants.LayerElementNodeName).Attributes();

            // Set layer properties.
            SetLayerProperties(layer, listOfAttributes);

            // Set marker properties.
            SetMarkerProperties(layer, listOfAttributes);

            return layer;
        }

        /// <summary>
        /// This function is used to set all layer properties of the layer.
        /// </summary>
        /// <param name="layer">
        /// Layer which is in focus.
        /// </param>
        /// <param name="listOfAttributes">
        /// Attributes of layer.
        /// </param>
        private static void SetLayerProperties(Layer layer, IEnumerable<XAttribute> listOfAttributes)
        {
            // If the Layer from Group SKY then the LAT is DEC and LON is RA respectively.
            // Process and update attributes in layer Details.
            foreach (XAttribute attribute in listOfAttributes)
            {
                switch (attribute.Name.LocalName)
                {
                    case Constants.LatColumnAttributeName:
                        if (layer.Group.IsPlanet())
                        {
                            layer.LatColumn = attribute.Value.AsInteger(Constants.DefaultColumnIndex);
                        }
                        else
                        {
                            layer.DecColumn = attribute.Value.AsInteger(Constants.DefaultColumnIndex);
                        }
                        break;
                    case Constants.LngColumnAttributeName:
                        if (layer.Group.IsPlanet())
                        {
                            layer.LngColumn = attribute.Value.AsInteger(Constants.DefaultColumnIndex);
                        }
                        else
                        {
                            layer.RAColumn = attribute.Value.AsInteger(Constants.DefaultColumnIndex);
                        }
                        break;
                    case Constants.GeometryColumnAttributeName:
                        layer.GeometryColumn = attribute.Value.AsInteger(Constants.DefaultColumnIndex);
                        break;
                    case Constants.ColorMapColumnAttributeName:
                        layer.ColorMapColumn = attribute.Value.AsInteger(Constants.DefaultColumnIndex);
                        break;
                    case Constants.AltColumnAttributeName:
                        layer.AltColumn = attribute.Value.AsInteger(Constants.DefaultColumnIndex);
                        break;
                    case Constants.StartDateColumnAttributeName:
                        layer.StartDateColumn = attribute.Value.AsInteger(Constants.DefaultColumnIndex);
                        break;
                    case Constants.EndDateColumnAttributeName:
                        layer.EndDateColumn = attribute.Value.AsInteger(Constants.DefaultColumnIndex);
                        break;
                    case Constants.SizeColumnAttributeName:
                        layer.SizeColumn = attribute.Value.AsInteger(Constants.DefaultColumnIndex);
                        break;
                    case Constants.NameColumnAttributeName:
                        layer.NameColumn = attribute.Value.AsInteger(Constants.DefaultColumnIndex);
                        break;
                    case Constants.NameAttribute:
                        layer.Name = attribute.Value;
                        break;
                    case Constants.XAxisColumnAttributeName:
                        layer.XAxis = attribute.Value.AsInteger(Constants.DefaultColumnIndex);
                        break;
                    case Constants.YAxisColumnAttributeName:
                        layer.YAxis = attribute.Value.AsInteger(Constants.DefaultColumnIndex);
                        break;
                    case Constants.ZAxisColumnAttributeName:
                        layer.ZAxis = attribute.Value.AsInteger(Constants.DefaultColumnIndex);
                        break;
                    case Constants.ReverseXAxisColumnAttributeName:
                        layer.ReverseXAxis = attribute.Value.AsBoolean(false);
                        break;
                    case Constants.ReverseYAxisColumnAttributeName:
                        layer.ReverseYAxis = attribute.Value.AsBoolean(false);
                        break;
                    case Constants.ReverseZAxisColumnAttributeName:
                        layer.ReverseZAxis = attribute.Value.AsBoolean(false);
                        break;
                    case Constants.VersionAttribute:
                        layer.Version = attribute.Value.AsInteger(Constants.DefaultLayerVersion);
                        break;
                }
            }
        }

        /// <summary>
        /// This function is used to set all marker properties of the layer.
        /// </summary>
        /// <param name="layer">
        /// Layer which is in focus.
        /// </param>
        /// <param name="listOfAttributes">
        /// Attributes of layer.
        /// </param>
        private static void SetMarkerProperties(Layer layer, IEnumerable<XAttribute> listOfAttributes)
        {
            // Process and update attributes in layer Details.
            foreach (XAttribute attribute in listOfAttributes)
            {
                switch (attribute.Name.LocalName)
                {
                    case Constants.DecayAttributeName:
                        layer.TimeDecay = attribute.Value.AsDouble(Constants.DefaultTimeDecay);
                        break;
                    case Constants.ScaleFactorAttributeName:
                        layer.ScaleFactor = attribute.Value.AsDouble(Constants.DefaultScaleFactor);
                        break;
                    case Constants.OpacityAttributeName:
                        layer.Opacity = attribute.Value.AsDouble(Constants.DefaultOpacity);
                        break;
                    case Constants.StartTimeAttributeName:
                        layer.StartTime = attribute.Value.AsDateTime(Constants.DefaultStartTime);
                        break;
                    case Constants.EndTimeAttributeName:
                        layer.EndTime = attribute.Value.AsDateTime(Constants.DefaultEndTime);
                        break;
                    case Constants.FadeSpanAttributeName:
                        layer.FadeSpan = attribute.Value.AsTimeSpan(Constants.DefaultFadeSpan);
                        break;
                    case Constants.ColorValueAttributeName:
                        layer.Color = string.IsNullOrEmpty(attribute.Value) ? Constants.DefaultColor : attribute.Value;
                        break;
                    case Constants.AltTypeAttributeName:
                        layer.AltType = attribute.Value.AsEnum<AltType>(AltType.Depth);
                        break;
                    case Constants.MarkerScaleAttributeName:
                        layer.MarkerScale = attribute.Value.AsEnum<ScaleRelativeType>(ScaleRelativeType.World);
                        break;
                    case Constants.CartesianScaleAttributeName:
                        // In case of non Spherical co-ordinates, value to be read from CartesianScale attribute.
                        if (layer.CoordinatesType != CoordinatesType.Spherical)
                        {
                            layer.AltUnit = attribute.Value.AsEnum<AltUnit>(AltUnit.Meters);
                        }
                        break;
                    case Constants.AltUnitAttributeName:
                        // In case of Spherical co-ordinates, value to be read from AltUnit attribute.
                        if (layer.CoordinatesType == CoordinatesType.Spherical)
                        {
                            layer.AltUnit = attribute.Value.AsEnum<AltUnit>(AltUnit.Meters);
                        }
                        break;
                    case Constants.RAUnitAttributeName:
                        layer.RAUnit = attribute.Value.AsEnum<AngleUnit>(AngleUnit.Hours);
                        break;
                    case Constants.PointScaleTypeAttributeName:
                        layer.PointScaleType = attribute.Value.AsEnum<ScaleType>(ScaleType.Power);
                        break;
                    case Constants.FadeTypeAttributeName:
                        layer.FadeType = attribute.Value.AsEnum<FadeType>(FadeType.None);
                        break;
                    case Constants.MarkerIndexAttributeName:
                        layer.MarkerIndex = attribute.Value.AsInteger(Constants.DefaultMarkerIndex);
                        break;
                    case Constants.MarkerTypeAttributeName:
                        layer.PlotType = attribute.Value.AsEnum<MarkerType>(MarkerType.Gaussian);
                        break;
                    case Constants.ShowFarSideAttributeName:
                        layer.ShowFarSide = attribute.Value.AsBoolean(false);
                        break;
                    case Constants.CoordinateTypeAttributeName:
                        layer.CoordinatesType = attribute.Value.AsEnum<CoordinatesType>(CoordinatesType.Spherical);
                        break;
                    case Constants.TimeSeriesAttributeName:
                        layer.HasTimeSeries = attribute.Value.AsBoolean(false);
                        break;
                }
            }
        }

        /// <summary>
        /// This function is used to retrieve payload for the updating the header.
        /// </summary>
        /// <param name="layer">Details of the layer.</param>
        /// <param name="isTimeSeriesRequired">If time series attribute is required to be set explicitly.</param>
        /// <returns>The payload of the header.</returns>
        private static string GetLayerProperties(Layer layer, bool isTimeSeriesRequired)
        {
            XAttribute cartesianScale = null;

            if (layer.CoordinatesType != CoordinatesType.Spherical)
            {
                cartesianScale = new XAttribute(Constants.CartesianScaleAttributeName, layer.AltUnit);
            }

            // If the Layer group is of type SKY (NOT PLANET) then LAT is DEC and Long is RA respectively.
            var layerElement = new XElement(
             Constants.LayerElementNodeName,
             new XAttribute(Constants.NameAttribute, layer.Name),
             new XAttribute(Constants.CoordinateTypeAttributeName, layer.CoordinatesType),
             new XAttribute(Constants.XAxisColumnAttributeName, layer.XAxis),
             new XAttribute(Constants.YAxisColumnAttributeName, layer.YAxis),
             new XAttribute(Constants.ZAxisColumnAttributeName, layer.ZAxis),
             new XAttribute(Constants.ReverseXAxisColumnAttributeName, layer.ReverseXAxis),
             new XAttribute(Constants.ReverseYAxisColumnAttributeName, layer.ReverseYAxis),
             new XAttribute(Constants.ReverseZAxisColumnAttributeName, layer.ReverseZAxis),
             new XAttribute(Constants.LatColumnAttributeName, layer.Group.IsPlanet() ? layer.LatColumn : layer.DecColumn),
             new XAttribute(Constants.LngColumnAttributeName, layer.Group.IsPlanet() ? layer.LngColumn : layer.RAColumn),
             new XAttribute(Constants.GeometryColumnAttributeName, layer.GeometryColumn),
             new XAttribute(Constants.ColorMapColumnAttributeName, layer.ColorMapColumn),
             new XAttribute(Constants.AltColumnAttributeName, layer.AltColumn),
             new XAttribute(Constants.StartDateColumnAttributeName, layer.StartDateColumn),
             new XAttribute(Constants.EndDateColumnAttributeName, layer.EndDateColumn),
             new XAttribute(Constants.SizeColumnAttributeName, layer.SizeColumn),
             new XAttribute(Constants.NameColumnAttributeName, layer.NameColumn),
             new XAttribute(Constants.DecayAttributeName, layer.TimeDecay),
             new XAttribute(Constants.ScaleFactorAttributeName, layer.ScaleFactor),
             new XAttribute(Constants.OpacityAttributeName, layer.Opacity),
             new XAttribute(Constants.StartTimeAttributeName, layer.StartTime.ToString()),
             new XAttribute(Constants.EndTimeAttributeName, layer.EndTime.ToString()),
             new XAttribute(Constants.FadeSpanAttributeName, layer.FadeSpan.ToString()),
             new XAttribute(Constants.ColorValueAttributeName, layer.Color),
             new XAttribute(Constants.AltTypeAttributeName, layer.AltType),
             new XAttribute(Constants.MarkerScaleAttributeName, layer.MarkerScale),
             new XAttribute(Constants.AltUnitAttributeName, layer.AltUnit),
             cartesianScale,
             new XAttribute(Constants.RAUnitAttributeName, layer.RAUnit),
             new XAttribute(Constants.PointScaleTypeAttributeName, layer.PointScaleType),
             new XAttribute(Constants.FadeTypeAttributeName, layer.FadeType),
             new XAttribute(Constants.MarkerTypeAttributeName, layer.PlotType),
             new XAttribute(Constants.MarkerIndexAttributeName, layer.MarkerIndex),
             new XAttribute(Constants.ShowFarSideAttributeName, layer.ShowFarSide));

            if (isTimeSeriesRequired)
            {
                layerElement.Add(new XAttribute(Constants.TimeSeriesAttributeName, layer.HasTimeSeries));
            }
            XElement root = new XElement(Constants.LCAPIElementName, layerElement);

            return root.ToString(SaveOptions.DisableFormatting);
        }

        /// <summary>
        /// Splits the given tab ('\t') delimited strings.
        /// </summary>
        /// <param name="data">String to be split.</param>
        /// <returns>List of strings</returns>
        private static List<string> SplitString(string data)
        {
            List<string> output = null;
            if (!string.IsNullOrWhiteSpace(data))
            {
                output = new List<string>(data.Split(new char[] { '\t' }));
            }
            else
            {
                // Send empty list of string instead of null.
                output = new List<string>();
            }

            return output;
        }

        /// <summary>
        /// Gets the data from WWT for the selected layer
        /// </summary>
        /// <param name="data">Layer data</param>
        /// <param name="cols">Number of columns</param>
        /// <returns>Object array with layer data from WWT</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1814:PreferJaggedArraysOverMultidimensional", MessageId = "Body", Justification = "We cannot use jagged array in this scenario because the excel Object model is designed to convert the value as [,]."), System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1814:PreferJaggedArraysOverMultidimensional", MessageId = "Return", Justification = "We cannot use jagged array in this scenario because the excel Object model is designed to convert the value as [,].")]
        private static object[,] GetAllLayerData(Collection<string[]> data, int cols)
        {
            object[,] dataValues = null;
            if (data != null && data.Count > 0)
            {
                int rows = data.Count;
                dataValues = new object[rows, cols];

                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < cols; col++)
                    {
                        try
                        {
                            if (data[row] != null && data[row].Length > 0)
                            {
                                dataValues[row, col] = data[row][col];
                            }
                        }
                        catch (IndexOutOfRangeException)
                        {
                            // If the data has comma separated values, then WWT returns tab separated data.
                            // the data will be send to excel as two separate columns
                            continue;
                        }
                    }
                }
            }
            return dataValues;
        }

        /// <summary>
        /// This function is used to check if the response is for valid Layer ID
        /// </summary>
        /// <param name="response">
        /// Response for the current operation.
        /// </param>
        /// <returns>
        /// True if the response contains valid Layer ID; Otherwise false.
        /// </returns>
        private static bool IsValidResponse(string response)
        {
            bool isValid = false;
            try
            {
                XElement root = XElement.Parse(response);
                isValid = string.Compare(root.Element(Constants.StatusAttribute).Value, Properties.Resources.ResponseSuccessfulText, StringComparison.OrdinalIgnoreCase) == 0;
            }
            catch (XmlException ex)
            {
                Logger.LogException(ex);

                // Ignore error.
            }

            return isValid;
        }

        /// <summary>
        /// This function is used to check if the response is for Timeout
        /// </summary>
        /// <param name="response">Response for the current operation.</param>
        /// <returns>True if the response is for Timeout; Otherwise false.</returns>
        private static bool IsTimeoutResponse(string response)
        {
            bool isTimeout = false;
            try
            {
                XElement root = XElement.Parse(response);
                isTimeout = string.Compare(root.Element(Constants.StatusAttribute).Value, Properties.Resources.TimeoutErrorText, StringComparison.OrdinalIgnoreCase) == 0;
            }
            catch (XmlException ex)
            {
                Logger.LogException(ex);

                // Ignore error.
            }

            return isTimeout;
        }

        /// <summary>
        /// This function is used to check if the response is valid version
        /// </summary>
        /// <param name="response">Response for the current operation</param>
        /// <returns>True if the response contains version</returns>
        private static bool IsValidVersion(string response)
        {
            bool isValid = false;
            try
            {
                XElement root = XElement.Parse(response);
                if (root.Element(Constants.VersionAttribute) != null && !string.IsNullOrEmpty(root.Element(Constants.VersionAttribute).Value))
                {
                    // Checks if the version of target machine is greater than or equal to base version
                    Version baseVersion = new Version(Constants.WWTBaseVersion);
                    Version targetVersion = new Version(root.Element(Constants.VersionAttribute).Value);
                    if (targetVersion >= baseVersion)
                    {
                        isValid = true;
                    }
                }
            }
            catch (XmlException ex)
            {
                Logger.LogException(ex);

                // Ignore error.
            }
            catch (ArgumentException ex)
            {
                Logger.LogException(ex);

                // Ignore error.
            }
            catch (OverflowException ex)
            {
                Logger.LogException(ex);

                // Ignore error.
            }
            catch (FormatException ex)
            {
                Logger.LogException(ex);

                // Ignore error.
            }
            return isValid;
        }

        /// <summary>
        /// This function is used to create groups for the specified Reference frame or layer groups.
        /// </summary>
        /// <param name="node">
        /// XML element of the group.
        /// </param>
        /// <param name="parent">
        /// Parent group.
        /// </param>
        /// <returns>
        /// Group object.
        /// </returns>
        private static Group GetGroup(XmlElement node, Group parent)
        {
            // Get Name.
            string name = node.Attributes[Constants.NameAttribute].Value;

            // Get Type.
            GroupType type = GroupType.None;
            if (node.Name == Constants.ReferenceFrameElementName)
            {
                type = GroupType.ReferenceFrame;
            }
            else if (node.Name == Constants.LayerGroupElementName)
            {
                type = GroupType.LayerGroup;
            }

            Group newGroup = new Group(name, type, parent);
            if (node.HasChildNodes)
            {
                foreach (XmlElement element in node.ChildNodes)
                {
                    if (element.Name == Constants.ReferenceFrameElementName || element.Name == Constants.LayerGroupElementName)
                    {
                        Group childGroup = GetGroup(element, newGroup);
                        if (childGroup != null)
                        {
                            newGroup.Children.Add(childGroup);
                        }
                    }
                    else if (element.Name == Constants.LayerElementNodeName && element.Attributes[Constants.IDAttribute] != null
                        && element.Attributes["Type"] != null && element.Attributes["Type"].Value.Equals("SpreadSheetLayer", StringComparison.OrdinalIgnoreCase))
                    {
                        newGroup.LayerIDs.Add(element.Attributes[Constants.IDAttribute].Value);
                    }
                }
            }

            return newGroup;
        }

        /// <summary>
        /// This function is used to recursively check if the layer ID is present in the group or not.
        /// </summary>
        /// <param name="layerId">
        /// ID of the layer.
        /// </param>
        /// <param name="group">
        /// Group in focus.
        /// </param>
        /// <returns>
        /// True, if the layer is present in group;Otherwise false.
        /// </returns>
        private static bool IsValidLayer(string layerId, Group group)
        {
            bool isValid = false;

            foreach (string layer in group.LayerIDs)
            {
                if (string.Compare(layer, layerId, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    isValid = true;
                    break;
                }
            }

            if (!isValid)
            {
                foreach (Group child in group.Children)
                {
                    if (IsValidLayer(layerId, child))
                    {
                        isValid = true;
                        break;
                    }
                }
            }

            return isValid;
        }

        /// <summary>
        /// This function is used to recursively check if the group is present in WWTgroup or not.
        /// </summary>
        /// <param name="group">
        /// Group in focus.
        /// </param>
        /// <param name="wwtGroup">
        /// List of WWT groups.
        /// </param>
        /// <returns>
        /// True, if the group is present in WWTgroup;Otherwise false.
        /// </returns>
        private static bool IsValidGroup(Group group, Group wwtGroup)
        {
            bool isValid = false;
            if (string.CompareOrdinal(group.Name, wwtGroup.Name) == 0 &&
                string.CompareOrdinal(group.Path, wwtGroup.Path) == 0)
            {
                isValid = true;
            }

            if (!isValid)
            {
                foreach (Group child in wwtGroup.Children)
                {
                    if (IsValidGroup(group, child))
                    {
                        isValid = true;
                        break;
                    }
                }
            }

            return isValid;
        }

        /// <summary>
        /// Get the attribute value from the node
        /// </summary>
        /// <param name="node">node object</param>
        /// <param name="attributeName">attribute Name</param>
        /// <returns>attribute value</returns>
        private static string GetAttributeValue(XmlNode node, string attributeName)
        {
            if (node.Attributes[attributeName] != null)
            {
                return node.Attributes[attributeName].Value;
            }

            return string.Empty;
        }

        /// <summary>
        /// Find the version of the Excel Add-In that is installed on the local machine
        /// </summary>
        /// <returns>The version number as a string</returns>
        private static string GetAddinInstalledVersion()
        {
            string productVersion = "0.0.0.0";

            try
            {
                // Since the installer is 32 bit, the registry entries for Excel Add-In will always go in to WOW6432Node registry.
                using (RegistryKey baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32))
                {
                    using (RegistryKey regKey = baseKey.OpenSubKey(Constants.UninstallRegistryPath))
                    {
                        if (regKey != null)
                        {
                            foreach (string subKeyName in regKey.GetSubKeyNames())
                            {
                                using (RegistryKey subKey = regKey.OpenSubKey(subKeyName))
                                {
                                    if (Convert.ToString(subKey.GetValue(Constants.DisplayNameProperty), CultureInfo.InvariantCulture).Equals(Properties.Resources.ProductNameShort))
                                    {
                                        productVersion = subKey.GetValue(Constants.DisplayVersionProperty).ToString();
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Security.SecurityException ex)
            {
                Logger.LogException(ex);
            }
            catch (ArgumentException ex)
            {
                Logger.LogException(ex);
            }
            catch (UnauthorizedAccessException ex)
            {
                Logger.LogException(ex);
            }
            catch (IOException ex)
            {
                Logger.LogException(ex);
            }

            return productVersion;
        }

        /// <summary>
        /// Sends asynchronous request to LCAPU for layer property change notification.
        /// </summary>
        /// <param name="command">Notify command</param>
        /// <param name="cancellationToken">Cancellation token to cancel the asynchronous request</param>
        /// <returns>Task having the string response</returns>
        private static Task<string> SendAsync(string command, CancellationToken cancellationToken)
        {
            TaskCompletionSource<string> taskCompletionSource = new TaskCompletionSource<string>();

            using (WebClient client = new WebClient())
            {
                // If user cancels the notification, then cancel the WebClient's asynchronous operations.
                cancellationToken.Register(() =>
                {
                    client.CancelAsync();
                });

                client.UploadStringAsync(new Uri(command), string.Empty);
                client.UploadStringCompleted += (obj, args) =>
                {
                    if (args.Cancelled == true || args.Error != null)
                    {
                        // Set the default error response in case of exception or cancellation.
                        string response = Constants.DefaultErrorResponse;
                        taskCompletionSource.TrySetResult(response);
                    }
                    else
                    {
                        string response = args.Result;
                        if (!string.IsNullOrEmpty(response))
                        {
                            try
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
                                }
                                else
                                {
                                    // Sanitize the response for hex characters
                                    response = Regex.Replace(response, Constants.HexCharacterPattern, string.Empty);
                                }
                            }
                            catch (XmlException exception)
                            {
                                Logger.LogException(exception);
                                response = Constants.DefaultErrorResponse;
                            }
                        }

                        taskCompletionSource.TrySetResult(response);
                    }
                };

                return taskCompletionSource.Task;
            }
        }

        #endregion
    }
}
