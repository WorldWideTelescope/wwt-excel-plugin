//-----------------------------------------------------------------------
// <copyright file="WorkflowController.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Media.Animation;
using Microsoft.Office.Interop.Excel;
using Microsoft.Research.Wwt.Excel.Addin.Properties;
using Microsoft.Research.Wwt.Excel.Common;
using System.Text;
using System.Xml.Linq;
using System.Net;
using System.IO;
using System.Xml;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// This class is responsible for all the interactions between the Excel and the WWT layer.
    /// </summary>
    internal class WorkflowController : IDisposable
    {
        private int columnCount;
        private int rowCount;
        private object[,] tempValue;

        /// <summary>
        /// Singleton instance
        /// </summary>
        private static WorkflowController instance;

        /// <summary>
        /// List of WorkbookMaps in excel application instance
        /// </summary>
        private List<WorkbookMap> workBookMaps;

        /// <summary>
        /// List of ViewpointMaps in excel application instance
        /// </summary>
        private List<ViewpointMap> viewpointMaps;

        /// <summary>
        /// current selected workbook map
        /// </summary>
        private WorkbookMap currentWorkbookMap;

        /// <summary>
        /// current selected Viewpoint map
        /// </summary>
        private ViewpointMap currentViewpointMap;

        /// <summary>
        /// LayerManagerPane control instance
        /// </summary>
        private LayerManagerPane layerManagerPaneInstance;

        /// <summary>
        /// Ribbon control instance
        /// </summary>
        private Ribbon ribbonInstance;

        /// <summary>
        /// UpdateManager instance
        /// </summary>
        private UpdateManager updateManager;

        /// <summary>
        /// ViewModel for the Download Updates button on the custom task pane
        /// </summary>
        private DownloadUpdatesViewModel downloadUpdatesViewModel;

        /// <summary>
        /// Manage viewpoint instance
        /// </summary>
        private ManageViewpoint manageViewpointInstance;

        /// <summary>
        /// layerDetailsViewModel instance
        /// </summary>
        private LayerDetailsViewModel layerDetailsViewModel;

        /// <summary>
        /// Store the most recently used workbook
        /// </summary>
        private Workbook mostRecentWorkbook;

        /// <summary>
        /// Store the most recently used worksheet
        /// </summary>
        private Worksheet mostRecentWorksheet;

        /// <summary>
        /// Stores the list of layers mapped to the sheet being deactivated
        /// </summary>
        private List<LayerMap> affectedLayers = new List<LayerMap>();

        /// <summary>
        /// Stores the affected named ranges when a cell's value changes.
        /// </summary>
        private Dictionary<string, string> affectedNamedRanges = new Dictionary<string, string>();

        /// <summary>
        /// Stores the handled named ranges from list of affected named ranges when a cell's value changes.
        /// </summary>
        private Dictionary<string, string> handledNamedRanges = new Dictionary<string, string>();

        /// <summary>
        /// Store the active worksheet name.
        /// </summary>
        private string worksheetName = null;

        /// <summary>
        /// Store the current worksheet instance.
        /// </summary>
        private Worksheet currentWorksheet = null;

        /// <summary>
        /// Stores Boolean which indicates whether AfterCalculate method is called or not.
        /// </summary>
        private bool afterCalculateCalled = false;

        /// <summary>
        /// Track whether Dispose has been called.
        /// </summary>
        private bool disposed = false;

        /// <summary>
        /// Prevents a default instance of the WorkflowController class from being created.
        /// </summary>
        private WorkflowController()
        {
            this.workBookMaps = new List<WorkbookMap>();
            this.viewpointMaps = new List<ViewpointMap>();
            this.InitializeUpdateManager();
            this.AttachUpdateManagerEventHandlers();
            this.AttachWorkbookEventHandlers();
            WorkflowController.LockObject = new object();

            // These objects may be null when this constructor is called by unit test cases.
            if (ThisAddIn.ExcelApplication != null && ThisAddIn.ExcelApplication.ActiveSheet != null)
            {
                currentWorksheet = (Worksheet)ThisAddIn.ExcelApplication.ActiveSheet;
                worksheetName = currentWorksheet.Name;
            }
        }

        /// <summary>
        /// Gets WorkflowController instance
        /// </summary>
        internal static WorkflowController Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new WorkflowController();
                }

                return instance;
            }
        }

        /// <summary>
        /// Gets the lock object used for locking while updating the Layer properties in background thread.
        /// </summary>
        internal static object LockObject { get; private set; }

        /// <summary>
        /// Gets or sets the last used group for the workbook.
        /// </summary>
        private static Group LastUsedGroup
        {
            get
            {
                Group lastUsed = null;
                if (!string.IsNullOrEmpty(Settings.Default.LastUsedGroup))
                {
                    lastUsed = lastUsed.Deserialize(Settings.Default.LastUsedGroup);
                }

                return lastUsed;
            }
            set
            {
                Settings.Default.LastUsedGroup = value.Serialize();
                Settings.Default.Save();
            }
        }

        #region Public Methods

        /// <summary>
        /// Part of IDisposable Interface
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion

        #region Internal Methods

        /// <summary>
        /// Checks if the selected layer map type is WWT or none. 
        /// </summary>
        /// <param name="layerMapType">Layer map type</param>
        /// <returns>True if the layer is local</returns>
        internal static bool IsLocalLayer(LayerMapType layerMapType)
        {
            return (layerMapType == LayerMapType.Local || layerMapType == LayerMapType.LocalInWWT);
        }

        /// <summary>
        /// Builds the reference frame dropdown
        /// </summary>
        internal void BuildReferenceFrameDropDown()
        {
            if (this.layerDetailsViewModel != null && WWTManager.IsValidMachine(Common.Globals.TargetMachine.MachineIP.ToString(), true))
            {
                ICollection<Group> groups = WWTManager.GetAllWWTGroups(true);
                if (groups != null && groups.Count > 0)
                {
                    GroupViewModel referenceGroup = new GroupViewModel();
                    referenceGroup.Name = string.Empty;
                    foreach (Group group in groups)
                    {
                        group.IsExpanded = true;
                        referenceGroup.ReferenceGroup.Add(group);
                    }
                    this.layerDetailsViewModel.ReferenceGroups.Add(referenceGroup);
                }
            }
        }

        /// <summary>
        /// This function is used to initialize the WorkflowController class.
        /// </summary>
        internal void Initialize(LayerManagerPane layerManagerPane, Ribbon ribbon)
        {
            if (ribbon != null)
            {
                this.layerManagerPaneInstance = layerManagerPane;
                this.ribbonInstance = ribbon;
                this.AttachRibbonEventHandlers();

                // Create WorkbookMap for Active workbook 
                // Check if Active workbook has a BookMap stored. If yes de-serialize it. Else create a new map
                this.currentWorkbookMap = ThisAddIn.ExcelApplication.ActiveWorkbook.GetWorkbookMap();
                this.currentViewpointMap = ThisAddIn.ExcelApplication.ActiveWorkbook.GetViewpointMap();

                // Add this to the workbook map list
                this.workBookMaps.Add(this.currentWorkbookMap);
                this.viewpointMaps.Add(this.currentViewpointMap);

                this.ribbonInstance.BuildViewpointMenu(this.currentViewpointMap.SerializablePerspective);
                this.BuildAndBindLayerDetailsViewModel();

                // check for updates
                this.updateManager.CheckForUpdates();

                //When launched from installer, Excel is already up and running before the addIn loads. Events are already fired!
                //Functionality of LayerTaskPaneController does not work as expected. So we explicitly set the WorkflowController to refer to the LayerManagerPane.
                if(LayerTaskPaneController.IsExcelInstanceSDI)
                    if (ThisAddIn.ExcelApplication.Workbooks.Count == 1)    //During Initialization if count is 1, workbook is already opened. Other scenarios count is 0 during Initialization
                        UpdateLayerManagerPaneInstance(LayerTaskPaneController.Instance.CurrentPaneHost.LayerManagerPane);
            }
        }

        /// <summary>
        /// Register for layer manager pane changed event. Pane changes when workbook changes (See - LayerTaskPaneController)
        /// </summary>
        internal void RegisterForLayerPaneChangedEvent()
        {
            LayerTaskPaneController.Instance.LayerPaneChangedEvent += new LayerPaneChanged(OnLayerPaneChangedEvent);
        }

        /// <summary>
        /// Gets the selected layer based on drop down properties
        /// </summary>
        /// <param name="selectedLayer">layerDropdown Properties</param>
        /// <returns>Layer map for selected layer</returns>
        internal LayerMap GetSelectedLayerMap(LayerMap selectedLayer)
        {
            LayerMap selectedlayermap = null;
            if (this.currentWorkbookMap != null && selectedLayer != null)
            {
                if (selectedLayer.MapType == LayerMapType.LocalInWWT)
                {
                    selectedlayermap = this.currentWorkbookMap.LocalLayerMaps.Find(layer => !string.IsNullOrEmpty(layer.LayerDetails.ID) && layer.LayerDetails.ID.Equals(selectedLayer.LayerDetails.ID, StringComparison.OrdinalIgnoreCase));
                }
                else if (selectedLayer.MapType == LayerMapType.Local)
                {
                    selectedlayermap = this.currentWorkbookMap.LocalLayerMaps.Find(layer => layer.RangeDisplayName.Equals(selectedLayer.RangeDisplayName, StringComparison.OrdinalIgnoreCase));
                }
                else if (selectedLayer.MapType == LayerMapType.WWT)
                {
                    selectedlayermap = this.currentWorkbookMap.WWTLayerMaps.Find(layer => layer.LayerDetails.ID.Equals(selectedLayer.LayerDetails.ID, StringComparison.OrdinalIgnoreCase));
                }
            }

            return selectedlayermap;
        }

        /// <summary>
        /// Rebuild group layer dropdown with the reference frame/
        /// layer group and layers
        /// </summary>
        internal void RebuildGroupLayerDropdown()
        {
            if (this.layerDetailsViewModel != null && this.currentWorkbookMap != null)
            {
                this.layerDetailsViewModel.Layers.Clear();

                // Add Select One as default
                LayerMapDropDownViewModel defaultLayerMapModel = new LayerMapDropDownViewModel();
                defaultLayerMapModel.Name = Resources.DefaultSelectedLayerName;
                defaultLayerMapModel.ID = "-1";
                this.layerDetailsViewModel.Layers.Add(defaultLayerMapModel);

                // Local group children for all local layers
                List<GroupChildren> localGroups = new List<GroupChildren>();
                LayerMapDropDownViewModel localLayerMapModel = new LayerMapDropDownViewModel();
                localLayerMapModel.Name = Properties.Resources.LocalLayerName;
                localLayerMapModel.ID = "0";
                this.currentWorkbookMap.LocalLayerMaps.ForEach(localLayer =>
                {
                    localGroups = localLayer.BuildGroupCollection(localGroups);
                });

                localGroups.ForEach(groupItem =>
                {
                    localLayerMapModel.GroupCollection.Add(groupItem);
                });
                this.layerDetailsViewModel.Layers.Add(localLayerMapModel);

                // WWT group children for all WWT layers
                LayerMapDropDownViewModel wwtLayerMapModel = new LayerMapDropDownViewModel();
                wwtLayerMapModel.Name = Properties.Resources.WWTLayerName;
                wwtLayerMapModel.ID = "1";
                List<GroupChildren> wwtGroups = new List<GroupChildren>();
                this.currentWorkbookMap.WWTLayerMaps.ForEach(wwtLayer =>
                {
                    wwtGroups = wwtLayer.BuildGroupCollection(wwtGroups);
                });

                wwtGroups.ForEach(groupItem =>
                {
                    wwtLayerMapModel.GroupCollection.Add(groupItem);
                });
                this.layerDetailsViewModel.Layers.Add(wwtLayerMapModel);
            }
        }

        /// <summary>
        /// Starts the show highlight animation animation
        /// </summary>
        internal void BeginShowHighlightAnimation()
        {
            if (this.layerManagerPaneInstance != null)
            {
                Storyboard callOutstoryboard = (Storyboard)this.layerManagerPaneInstance.FindResource(Common.Constants.ShowHighlightAnimation);
                callOutstoryboard.Begin();
            }
        }

        /// <summary>
        /// Starts the hide highlight animation animation
        /// </summary>
        internal void BeginHideHighlightAnimation()
        {
            if (this.layerManagerPaneInstance != null)
            {
                Storyboard callOutstoryboard = (Storyboard)this.layerManagerPaneInstance.FindResource(Common.Constants.HideHighlightAnimation);
                callOutstoryboard.Begin();
            }
        }

        /// <summary>
        /// Gets the workbook map to which the given layer map belongs to.
        /// </summary>
        /// <param name="currentLayerMap">LayerMap for which WorkbookMap has to be fetched</param>
        /// <returns>WorkbookMap instance</returns>
        internal WorkbookMap GetWorkbookMapForLayerMap(LayerMap currentLayerMap)
        {
            WorkbookMap layerWorkbookMap = null;

            foreach (WorkbookMap workbookMap in this.workBookMaps)
            {
                if (workbookMap.AllLayerMaps.Where(layerMap => layerMap == currentLayerMap).FirstOrDefault() != null)
                {
                    layerWorkbookMap = workbookMap;
                    break;
                }
            }

            return layerWorkbookMap;
        }

        /// <summary>
        /// Builds and binds latest current layer details view model to layer details pane
        /// </summary>
        /// <param name="rebuildReferenceFrameDropDown">
        /// Optional parameter which tells whether reference frame dropdown needs to be rebuilt or not. If reference frame dropdown needs to be rebuilt, 
        /// an additional call to WWT API will be made. Default value is true (reference frame will be rebuilt).
        /// </param>
        /// <param name="isCallOutRequired">
        /// Optional parameter which tells whether call out needs to shown. This is to avoid the show of callout every time the layer dropdown
        /// is clicked.
        /// </param>
        internal void BuildAndBindLayerDetailsViewModel(bool rebuildReferenceFrameDropDown = true, bool isCallOutRequired = true)
        {
            lock (WorkflowController.LockObject)
            {
                //Set the rendering timeout check box initial state to checked. 
                bool renderingTimeoutCheckBoxState = true;
                //Also retain state each time the model is being recreated! (Overrides default except first time)
                if (this.layerDetailsViewModel != null)
                    renderingTimeoutCheckBoxState = this.layerDetailsViewModel.IsRenderingTimeoutAlertShown;

                this.layerDetailsViewModel = new LayerDetailsViewModel();
                this.layerDetailsViewModel.IsRenderingTimeoutAlertShown = renderingTimeoutCheckBoxState;
                
                LayerDetailsViewModel.IsCallOutRequired = isCallOutRequired;

                // Directly binding to DataContext is resulting in not considering flag value for IsPropertyChangedFromCode for relay commands
                this.layerDetailsViewModel = this.layerDetailsViewModel.BuildLayerDetailsViewModel(this.currentWorkbookMap, rebuildReferenceFrameDropDown);
                SetGetLayerDataDisplayName(this.currentWorkbookMap.SelectedLayerMap);

                if (this.layerManagerPaneInstance != null)
                {
                    LayerDetailsViewModel.IsPropertyChangedFromCode = true;
                    this.layerManagerPaneInstance.DataContext = this.layerDetailsViewModel;
                    this.layerDetailsViewModel.DownloadUpdatesViewModelInstance = this.downloadUpdatesViewModel;
                    LayerDetailsViewModel.IsPropertyChangedFromCode = false;

                    this.AttachCustomTaskEventHandlers();
                    this.layerManagerPaneInstance.UpdateLayout();
                }
            }
        }

        internal void CreateLayerMap()
        {
            // Workflow tasks for Visualize selection
            Range selectedRange = ThisAddIn.ExcelApplication.Selection as Range;

            if (selectedRange != null)
            {
                bool isWWTRunning = WWTManager.IsValidMachine(Common.Globals.TargetMachine.MachineIP.ToString(), true);

                // Gets the layer map for selected range if it already exists and selects the layer 
                LayerMap layerMap = GetCurrentRangeLayer(selectedRange);
                if (layerMap != null)
                {
                    this.currentWorkbookMap.SelectedLayerMap = layerMap;
                    this.currentWorkbookMap.SelectedLayerMap.VisualizeClickTime = DateTime.Now;
                    this.currentWorkbookMap.SelectedLayerMap.IsVisualizeClicked = true;
                }
                else
                {
                    string selectionRangeName = ThisAddIn.ExcelApplication.ActiveWorkbook.GetSelectionRangeName();
                    Name namedRange = ThisAddIn.ExcelApplication.ActiveWorkbook.CreateNamedRange(selectionRangeName, selectedRange);
                    if (namedRange != null)
                    {
                        LayerMap newLayerMap = new LayerMap(namedRange);

                        newLayerMap.LayerDetails.Group = GetLastUsedGroup(newLayerMap, isWWTRunning);

                        // Set the scale type based on the group the layer belongs to.
                        newLayerMap.LayerDetails.PointScaleType = newLayerMap.LayerDetails.Group.IsPlanet() ? ScaleType.Power : ScaleType.StellarMagnitude;

                        // Add the layer map for newly created named range to the list of LayerMap.
                        this.currentWorkbookMap.AllLayerMaps.Add(newLayerMap);

                        // Set the current layer map to the newly created LayerMap.
                        this.currentWorkbookMap.SelectedLayerMap = newLayerMap;

                        this.currentWorkbookMap.SelectedLayerMap.VisualizeClickTime = DateTime.Now;
                        this.currentWorkbookMap.SelectedLayerMap.IsVisualizeClicked = true;
                        ThisAddIn.ExcelApplication.ActiveWorkbook.SaveWorkbookMap(this.workBookMaps);
                    }
                }
                this.BuildAndBindLayerDetailsViewModel(isWWTRunning);
                if (this.ribbonInstance != null)
                {
                    this.ribbonInstance.ViewCustomTaskPane(true);

                    // Set focus on custom task pane and start animation for call out 
                    this.ribbonInstance.SetFocusCustomTaskPane();
                }
                this.BeginCalloutAnimation();

                // reset the default tab to Map columns tab (Index: 0)
                LayerDetailsViewModel.IsPropertyChangedFromCode = true;
                this.layerDetailsViewModel.SelectedTabIndex = 0;
                LayerDetailsViewModel.IsPropertyChangedFromCode = false;
            }
        }

        internal void GenerateWWTColumns(UpdateDataModel input)
        {
            object[,] updatedData = UpdatedWWTData(input);

            if (updatedData != null)
            {
                ThisAddIn.ExcelApplication.DisplayAlerts = false;

                string newSheetname = "Updated - " + (this.currentWorksheet != null ? this.currentWorksheet.Name : string.Empty);

                // Checking existing sheets for the FetchClimate sheet, if exists deleting the sheet.
                if (ThisAddIn.ExcelApplication.ActiveWorkbook.Sheets.Count > 0)
                {
                    foreach (_Worksheet sheet in ThisAddIn.ExcelApplication.ActiveWorkbook.Sheets)
                    {
                        if (sheet.Name == newSheetname)
                        {
                            sheet.Delete();
                            break;
                        }
                    }
                }

                ThisAddIn.ExcelApplication.DisplayAlerts = true;

                // Creating new sheet
                ThisAddIn.ExcelApplication.ActiveWorkbook.Sheets.Add();
                _Worksheet workSheet = (_Worksheet)ThisAddIn.ExcelApplication.ActiveWorkbook.ActiveSheet;
                workSheet.Name = newSheetname;

                int rowCount = updatedData.GetLength(0);

                Office.Interop.Excel.Range excelRange = workSheet.get_Range("A1", ExcelColumnName(updatedData.GetLength(1)) + rowCount.ToString());

                // Setting data to the excel.
                excelRange.SetValue(updatedData);

                Office.Interop.Excel.Range selectRange = workSheet.get_Range(ExcelColumnName(updatedData.GetLength(1) - 2) + "1", ExcelColumnName(updatedData.GetLength(1)) + rowCount.ToString());

                selectRange.Select();

                WorkflowController.Instance.CreateLayerMap();
            }
        }

        internal void GenerateHuricaneColumns(UpdateDataModel input)
        {
            object[,] updatedData = null;
            updatedData = UpdatedHuricaneData(input, updatedData);

            if (updatedData != null)
            {
                ThisAddIn.ExcelApplication.DisplayAlerts = false;

                string newSheetname = "Updated - " + (this.currentWorksheet != null ? this.currentWorksheet.Name : string.Empty);

                // Checking existing sheets for the FetchClimate sheet, if exists deleting the sheet.
                if (ThisAddIn.ExcelApplication.ActiveWorkbook.Sheets.Count > 0)
                {
                    foreach (_Worksheet sheet in ThisAddIn.ExcelApplication.ActiveWorkbook.Sheets)
                    {
                        if (sheet.Name == newSheetname)
                        {
                            sheet.Delete();
                            break;
                        }
                    }
                }

                ThisAddIn.ExcelApplication.DisplayAlerts = true;

                // Creating new sheet
                ThisAddIn.ExcelApplication.ActiveWorkbook.Sheets.Add();
                _Worksheet workSheet = (_Worksheet)ThisAddIn.ExcelApplication.ActiveWorkbook.ActiveSheet;
                workSheet.Name = newSheetname;

                int rowCount = updatedData.GetLength(0);

                Office.Interop.Excel.Range excelRange = workSheet.get_Range("A1", ExcelColumnName(updatedData.GetLength(1)) + rowCount.ToString());

                // Setting data to the excel.
                excelRange.SetValue(updatedData);

                excelRange.Select();

                WorkflowController.Instance.CreateLayerMap();
            }
        }

        internal void InsertFetchClimateData(double latMin, double latMax, double longMin, double longMax, double dlat, double dlong)
        {
            // Code to fetch the climate values from the Fetch climate API
            List<FetchClimateOutputModel> lstFetchClimateValues = FetchClimateAPIUtility.GetPrecipitationAndTemp(latMin, latMax, longMin, longMax, dlat, dlong);

            if (lstFetchClimateValues != null && lstFetchClimateValues.Count > 0)
            {
                object[,] dataValues = new object[lstFetchClimateValues.Count + 1, 9];

                // Assigning headers.
                dataValues[0, 0] = "MinLa";
                dataValues[0, 1] = "MaxLa";
                dataValues[0, 2] = "MinLn";
                dataValues[0, 3] = "MaxLn";
                dataValues[0, 4] = "Geometry";
                dataValues[0, 5] = "Precipitation";
                dataValues[0, 6] = "Temperature";
                dataValues[0, 7] = "Altitude";
                dataValues[0, 8] = "Color";

                // Assigning data to bind to the excel.
                for (int i = 1; i <= lstFetchClimateValues.Count; i++)
                {
                    dataValues[i, 0] = lstFetchClimateValues[i - 1].MinLatitude.ToString();
                    dataValues[i, 1] = lstFetchClimateValues[i - 1].MaxLatitude.ToString();
                    dataValues[i, 2] = lstFetchClimateValues[i - 1].MinLongitude.ToString();
                    dataValues[i, 3] = lstFetchClimateValues[i - 1].MaxLongitude.ToString();
                    dataValues[i, 4] = lstFetchClimateValues[i - 1].Geometry;
                    dataValues[i, 5] = lstFetchClimateValues[i - 1].Precipitation.ToString();
                    dataValues[i, 6] = lstFetchClimateValues[i - 1].Temperature.ToString();
                    dataValues[i, 7] = lstFetchClimateValues[i - 1].Altitude.ToString();
                    dataValues[i, 8] = lstFetchClimateValues[i - 1].Color;
                }

                ThisAddIn.ExcelApplication.DisplayAlerts = false;

                // Checking existing sheets for the FetchClimate sheet, if exists deleting the sheet.
                if (ThisAddIn.ExcelApplication.ActiveWorkbook.Sheets.Count > 0)
                {
                    foreach (_Worksheet sheet in ThisAddIn.ExcelApplication.ActiveWorkbook.Sheets)
                    {
                        if (sheet.Name == "FetchClimate")
                        {
                            sheet.Delete();
                            break;
                        }
                    }
                }

                ThisAddIn.ExcelApplication.DisplayAlerts = true;

                // Creating new sheet
                ThisAddIn.ExcelApplication.ActiveWorkbook.Sheets.Add();
                _Worksheet workSheet = (_Worksheet)ThisAddIn.ExcelApplication.ActiveWorkbook.ActiveSheet;
                workSheet.Name = "FetchClimate";

                int rowCount = lstFetchClimateValues.Count + 1;

                Office.Interop.Excel.Range excelRange = workSheet.get_Range("A1", "I" + rowCount);

                // Setting data to the excel.
                excelRange.SetValue(dataValues);

                excelRange.Select();

                WorkflowController.Instance.CreateLayerMap();
            }
        }

        internal void ShowAGUDemo()
        {
            double latMin = 34;
            double latMax = 36;
            double longMin = 69;
            double longMax = 71;
            double dlat = 0.04;
            double dlong = 0.04;

            bool valueUpdated = false;

            try
            {
                if (ThisAddIn.ExcelApplication.ActiveWorkbook != null)
                {
                    _Worksheet defaultWorksheet = null;
                    foreach (_Worksheet sheet in ThisAddIn.ExcelApplication.ActiveWorkbook.Sheets)
                    {
                        if (string.Compare(sheet.Name, "default", true) == 0)
                        {
                            defaultWorksheet = sheet;
                            break;
                        }
                    }

                    if (defaultWorksheet != null)
                    {
                        string A1 = defaultWorksheet.get_Range("A1").Value == null ? string.Empty : defaultWorksheet.get_Range("A1").Value.ToString(); // Delta Lat
                        string A2 = defaultWorksheet.get_Range("A2").Value == null ? string.Empty : defaultWorksheet.get_Range("A2").Value.ToString(); // Delta Lon
                        string A3 = defaultWorksheet.get_Range("A3").Value == null ? string.Empty : defaultWorksheet.get_Range("A3").Value.ToString(); // Center Lat
                        string A4 = defaultWorksheet.get_Range("A4").Value == null ? string.Empty : defaultWorksheet.get_Range("A4").Value.ToString(); // Center Lon
                        string A5 = defaultWorksheet.get_Range("A5").Value == null ? string.Empty : defaultWorksheet.get_Range("A5").Value.ToString(); // Lat cell count
                        string A6 = defaultWorksheet.get_Range("A6").Value == null ? string.Empty : defaultWorksheet.get_Range("A6").Value.ToString(); // Lon cell count

                        double latitude = 35.0;
                        if (!double.TryParse(A3, out latitude))
                        {
                            latitude = 35.0;
                        }

                        double longitude = 70.0;
                        if (!double.TryParse(A4, out longitude))
                        {
                            longitude = 70.0;
                        }

                        double latitudeCellCount = 2.0;
                        if (!double.TryParse(A5, out latitudeCellCount))
                        {
                            latitudeCellCount = 2.0;
                        }

                        double longitudeCellCount = 2.0;
                        if (!double.TryParse(A6, out longitudeCellCount))
                        {
                            longitudeCellCount = 2.0;
                        }

                        latMin = latitude - latitudeCellCount / 2; // Dividing by 2 to make sure evenly spread out for the given center lat and lon
                        latMax = latitude + latitudeCellCount / 2;
                        longMin = longitude - longitudeCellCount / 2;
                        longMax = longitude + longitudeCellCount / 2;

                        if (!double.TryParse(A1, out dlat))
                        {
                            dlat = 0.04;
                        }

                        if (!double.TryParse(A2, out dlong))
                        {
                            dlong = 0.04;
                        }

                        valueUpdated = true;
                    }
                }
            }
            catch { valueUpdated = false; } // Ignore any exceptions

            if (!valueUpdated)
            {
                // Get values from view point.
                Perspective perspective = WWTManager.GetCameraView();
                if (perspective != null)
                {
                    double latitude = 35.0;
                    if (!double.TryParse(perspective.Latitude, out latitude))
                    {
                        latitude = 35.0;
                    }

                    double longitude = 70.0;
                    if (!double.TryParse(perspective.Longitude, out longitude))
                    {
                        longitude = 70.0;
                    }

                    latMin = latitude - 1;
                    latMax = latitude + 1;
                    longMin = longitude - 1;
                    longMax = longitude + 1;
                    dlat = 0.04;
                    dlong = 0.04;
                }
            }

            // Get data from Fetch climate API and then create new layer
            InsertFetchClimateData(latMin, latMax, longMin, longMax, dlat, dlong);

            // Push current layer data to WWT.
            ViewInWWT();
        }

        internal void UpdateCurrentLocation()
        {
            // Get values from view point.
            Perspective perspective = WWTManager.GetCameraView();
            if (perspective != null)
            {
                double latitude = 35.0;
                if (!double.TryParse(perspective.Latitude, out latitude))
                {
                    latitude = 35.0;
                }

                double longitude = 70.0;
                if (!double.TryParse(perspective.Longitude, out longitude))
                {
                    longitude = 70.0;
                }

                var activeSheet = GetActiveWorksheet();

                if (activeSheet != null)
                {
                    var rowCount = activeSheet.UsedRange.Rows.Count;

                    if (rowCount == 1)
                    {
                        activeSheet.get_Range("A" + rowCount).Value2 = "Latitude";
                        activeSheet.get_Range("B" + rowCount).Value2 = "Longitude";
                    }

                    activeSheet.get_Range("A" + (rowCount + 1)).Value2 = perspective.Latitude;
                    activeSheet.get_Range("B" + (rowCount + 1)).Value2 = perspective.Longitude;
                }
            }
        }

        public static string ExcelColumnName(int column)
        {
            column--;
            if (column >= 0 && column < 26)
                return ((char)('A' + column)).ToString();
            else if (column > 25)
                return ExcelColumnName(column / 26) + ExcelColumnName(column % 26 + 1);
            else
                throw new Exception("Invalid Column #" + (column + 1).ToString());
        }

        internal Collection<string> GetColumns()
        {
            return GetActiveWorksheet().UsedRange.GetHeader();
        }

        internal void GetStationData(StationViewModel selectedViewModel)
        {
            StationDataModel model = selectedViewModel.StationData;

            StringBuilder sb = BuildQuery(model);

            HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(new Uri(sb.ToString()));
            myRequest.UserAgent = "Mozilla/4.0+";

            XDocument ResponseText;
            using (HttpWebResponse response = (HttpWebResponse)myRequest.GetResponse())
            {
                using (Stream responseStream = response.GetResponseStream())
                {
                    ResponseText = XDocument.Load(responseStream);
                }
            }

            string data = GetDatafromFeed(ResponseText.ToString());

            _Worksheet workSheet = GetActiveWorksheet();

            // Gets the range from the excel for data row and columns
            Range currentRange = workSheet.GetRange(Globals.ThisAddIn.Application.ActiveCell, rowCount, columnCount);
            if (currentRange != null)
            {
                currentRange.Select();
                currentRange.Value2 = tempValue;
            }

            WorkflowController.Instance.CreateLayerMap();
        }

        internal void GetEarthquakeData(StationViewModel selectedViewModel)
        {
            EarthquakeDataModel model = selectedViewModel.EarthquakeData;

            StringBuilder sb = BuildEarthquakeQuery(model);

            HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(new Uri(sb.ToString()));
            myRequest.UserAgent = "Mozilla/4.0+";

            XDocument ResponseText;
            using (HttpWebResponse response = (HttpWebResponse)myRequest.GetResponse())
            {

                using (Stream responseStream = response.GetResponseStream())
                {
                    ResponseText = XDocument.Load(responseStream);
                }
            }

            string data = GetEarthquakeDatafromFeed(ResponseText.ToString());

            _Worksheet workSheet = GetActiveWorksheet();

            // Gets the range from the excel for data row and columns
            Range currentRange = workSheet.GetRange(Globals.ThisAddIn.Application.ActiveCell, rowCount, columnCount);
            if (currentRange != null)
            {
                currentRange.Select();
                currentRange.Value2 = tempValue;
            }

            WorkflowController.Instance.CreateLayerMap();
        }

        #endregion Internal Methods

        #region Protected Methods

        /// <summary>
        /// Part of IDisposable Interface
        /// </summary>
        /// <param name="disposing">True if called from code</param>
        protected virtual void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!this.disposed)
            {
                // If disposing equals true, dispose all managed
                // and unmanaged resources.
                if (disposing)
                {
                    this.updateManager.Dispose();
                }

                // Note disposing has been done.
                disposed = true;
            }
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Updates the currently refered to layer pane
        /// </summary>
        /// <param name="layerManagerPane"></param>
        private void UpdateLayerManagerPaneInstance(LayerManagerPane layerManagerPane)
        {
            this.layerManagerPaneInstance = layerManagerPane;
            this.layerManagerPaneInstance.DataContext = this.layerDetailsViewModel;
        }

        /// <summary>
        /// This function is used to update header details in object models.
        /// </summary>
        /// <param name="selectedlayer">
        /// Updated layer.
        /// </param>
        private static void UpdateHeader(LayerMap selectedlayer)
        {
            // Integration with WWT : Update the Layer details in WWT for the changes 
            if (selectedlayer.CanUpdateWWT())
            {
                // This is used only for Layer Property update notifications. Since the notifications are handled in background thread,
                // CodeUpdate cannot be set the false here (main thread). It needs to be reset only by background thread once the 
                // notification is handled by the background thread.
                selectedlayer.IsPropertyChangedFromCode = true;

                // Update the header details.
                if (!WWTManager.UpdateLayer(selectedlayer.LayerDetails, true, false))
                {
                    selectedlayer.IsNotInSync = true;
                }
            }
        }

        /// <summary>
        /// This function is used to update data in WWT.
        /// </summary>
        /// <param name="selectedRange">
        /// Updated range.
        /// </param>
        /// <param name="selectedlayer">
        /// Updated layer.
        /// </param>
        private static void UpdateData(Range selectedRange, LayerMap selectedlayer)
        {
            // Update Data based on the Header
            // 1. Set layer properties dependent on mapping. 
            // 2. Purge the existing layer data in WWT and then push new data to WWT
            if (selectedlayer.CanUpdateWWT())
            {
                string[] data = selectedRange.GetData();

                if (!WWTManager.UploadDataInWWT(selectedlayer.LayerDetails.ID, data, true))
                {
                    selectedlayer.IsNotInSync = true;
                }
            }
        }

        /// <summary>
        /// This function is used to update the header properties and data in WWT.
        /// </summary>
        /// <param name="selectedLayerMap">
        /// Selected layer map.
        /// </param>
        /// <returns>
        /// True if the update is successfully completed;Otherwise false.
        /// </returns>
        private static bool UpdateWWT(LayerMap selectedLayerMap)
        {
            bool hasUpdated = false;

            if (selectedLayerMap.RangeName.IsValid())
            {
                // Get data from selected range.
                string[] data = selectedLayerMap.RangeName.RefersToRange.GetData();

                // This is used only for Layer Property update notifications. Since the notifications are handled in background thread,
                // CodeUpdate cannot be set the false here (main thread). It needs to be reset only by background thread once the 
                // notification is handled by the background thread.
                selectedLayerMap.IsPropertyChangedFromCode = true;

                // Update the header details.
                // Time series is set explicitly to false only on creation of layer because
                // on activate layer the time series is set to true while creating the layer even though
                // the value for time series is not send through LCAPI.
                if (WWTManager.UpdateLayer(selectedLayerMap.LayerDetails, true, true))
                {
                    var lookAt = selectedLayerMap.GetLookAt();
                    WWTManager.SetMode(lookAt);

                    // Upload data in WWT.
                    if (WWTManager.UploadDataInWWT(selectedLayerMap.LayerDetails.ID, data, true))
                    {
                        hasUpdated = true;
                    }
                }
            }

            return hasUpdated;
        }

        /// <summary>
        /// This function is used to create WWT layer if not exist.
        /// </summary>
        /// <param name="selectedLayerMap">
        /// Selected layer map.
        /// </param>
        /// <returns>True, if the creation is successful; otherwise false.</returns>
        private static bool CreateIfNotExist(LayerMap selectedLayerMap)
        {
            bool success = true;
            if (string.IsNullOrEmpty(selectedLayerMap.LayerDetails.ID) && selectedLayerMap.MapType == LayerMapType.Local)
            {
                success = CreateLayerInWWT(selectedLayerMap);
            }
            else if (!string.IsNullOrEmpty(selectedLayerMap.LayerDetails.ID) && selectedLayerMap.MapType == LayerMapType.LocalInWWT)
            {
                // Check if the Layer is present in WWT or not.
                if (!WWTManager.IsValidLayer(selectedLayerMap.LayerDetails.ID))
                {
                    success = CreateLayerInWWT(selectedLayerMap);
                }
            }

            return success;
        }

        /// <summary>
        /// This function is used to create Layer in WWT.
        /// </summary>
        /// <param name="selectedLayerMap">
        /// Selected layer map.
        /// </param>
        /// <returns>
        /// True, if the creation is successful; otherwise false.
        /// </returns>
        private static bool CreateLayerInWWT(LayerMap selectedLayerMap)
        {
            bool success = true;

            // Get Header Data
            string headerData = string.Join("\t", selectedLayerMap.HeaderRowData);

            // Set the version as 0. Serialized layers may have different version number which needs to be reset to 0 so that notification will work fine.
            selectedLayerMap.LayerDetails.Version = 0;
            ICollection<Group> wwtGroups = WWTManager.GetAllWWTGroups(true);

            if (!WWTManager.IsValidGroup(selectedLayerMap.LayerDetails.Group, wwtGroups))
            {
                // If the layer group is not present in WWT, then create the layer group before creating the layer.
                success = CreateGroupInWWT(selectedLayerMap);
            }

            if (success)
            {
                // if the layer is not present in WWT. Create the layer with the required Details.
                // Create Named Range. && Update layer id in the selected layer details.
                selectedLayerMap.LayerDetails.ID = WWTManager.CreateLayer(
                    selectedLayerMap.LayerDetails.Name,
                    selectedLayerMap.LayerDetails.Group.Name,
                    headerData);
                selectedLayerMap.MapType = LayerMapType.LocalInWWT;

                success = true;
            }

            return success;
        }

        /// <summary>
        /// This function is used to create layer group in WWT.
        /// </summary>
        /// <param name="selectedLayerMap">
        /// Selected layer map.
        /// </param>
        /// <returns>True, if the creation is successful; otherwise false.</returns>
        private static bool CreateGroupInWWT(LayerMap selectedLayerMap)
        {
            bool success = false;
            Group group = selectedLayerMap.LayerDetails.Group;
            if (group.GroupType == GroupType.LayerGroup)
            {
                try
                {
                    WWTManager.CreateLayerGroup(group.Name, group.Parent != null ? group.Parent.Name : string.Empty);
                    success = true;
                }
                catch (CustomException)
                {
                    // Could not create layer group in WWT.
                    Ribbon.ShowError(Properties.Resources.LayerGroupCreationError);
                    success = false;
                }
            }
            else
            {
                // Reference Frame cannot be created in WWT.
                Ribbon.ShowError(Properties.Resources.ReferenceFrameCreationError);
                success = false;
            }

            return success;
        }

        /// <summary>
        /// Gets the row/column difference
        /// </summary>
        /// <param name="rangeLength">Total range length</param>
        /// <param name="dataLength">Layer data length</param>
        /// <returns>Difference between range and data length</returns>
        private static int GetRangeDifference(int rangeLength, int dataLength)
        {
            int difference = 0;
            if (dataLength > rangeLength)
            {
                difference = dataLength - rangeLength;
            }
            else
            {
                difference = rangeLength - dataLength;
            }
            return difference;
        }

        /// <summary>
        /// Takes user to the viewpoint
        /// </summary>
        /// <param name="perspective">perspective object</param>
        private static void GotoViewpoint(Perspective perspective)
        {
            try
            {
                Utility.IsWWTInstalled();
                WWTManager.IsValidMachine(Common.Globals.TargetMachine.MachineIP.ToString(), false);
                WWTManager.SetCameraView(perspective);
                if (TargetMachine.DefaultIP.ToString() == Common.Globals.TargetMachine.MachineIP.ToString())
                {
                    Utility.ShowWWT();
                }
            }
            catch (CustomException ex)
            {
                Ribbon.ShowError(ex.HasCustomMessage ? ex.Message : Resources.LayerOperationError);
            }
        }

        /// <summary>
        /// Takes user to the viewpoint
        /// </summary>
        /// <param name="perspective">perspective object</param>
        private static void GotoViewpointFromData(Perspective perspective)
        {
            try
            {
                Utility.IsWWTInstalled();
                WWTManager.IsValidMachine(Common.Globals.TargetMachine.MachineIP.ToString(), false);
                WWTManager.SetCameraView(perspective);
                if (TargetMachine.DefaultIP.ToString() == Common.Globals.TargetMachine.MachineIP.ToString())
                {
                    Utility.ShowWWT();
                }
            }
            catch (CustomException ex)
            {
                // Earth or solar system
                if (!perspective.LookAt.Equals(Common.Constants.SkyLookAt, StringComparison.OrdinalIgnoreCase))
                {
                    // If Lat and Lon values were not mapped or invalid, show error
                    float validValue;
                    if (!float.TryParse(perspective.Latitude, out validValue) || !float.TryParse(perspective.Longitude, out validValue))
                    {
                        Ribbon.ShowError(Resources.GotoViewpointfromInvalidLatLonError);
                    }
                }
                else if (perspective.HasRADec)
                {
                    // If RA and Dec values were not mapped or invalid, show error.
                    float validValue;
                    if (!float.TryParse(perspective.RightAscention, out validValue) || !float.TryParse(perspective.Declination, out validValue))
                    {
                        Ribbon.ShowError(Resources.GotoViewpointfromInvalidRaDecError);
                    }
                }
                else
                {
                    Ribbon.ShowError(ex.HasCustomMessage ? ex.Message : Resources.LayerOperationError);
                }
            }
        }

        /// <summary>
        /// Get perspective based on the layer mapping and the row data
        /// </summary>
        /// <param name="layerMap">layer map object</param>
        /// <param name="rowData">row data collection</param>
        /// <returns>Perspective object</returns>
        private static Perspective GetPerspectiveFromLayerRowData(LayerMap layerMap, Collection<string> rowData)
        {
            Perspective perspective = null;
            if (layerMap.LayerDetails.Group != null)
            {
                var referenceFrame = layerMap.LayerDetails.Group.GetReferenceFrame();
                var referenceFramePath = layerMap.LayerDetails.Group.Path;
                if (referenceFramePath.StartsWith(Common.Constants.SkyFramePath, StringComparison.OrdinalIgnoreCase))
                {
                    perspective = new Perspective(Common.Constants.SkyLookAt, referenceFrame, true, Common.Constants.LatitudeDefaultValue, Common.Constants.LongitudeDefaultValue, Common.Constants.ZoomDefaultValue, Common.Constants.RotationDefaultValue, Common.Constants.LookAngleDefaultValue, DateTime.Now.ToString(), Common.Constants.TimeRateDefaultValue, Common.Constants.SkyZoomTextDefaultValue, string.Empty);
                    float hours = 0;
                    string rowValueForRA = rowData[layerMap.MappedColumnType.IndexOf(ColumnType.RA)];
                    if (layerMap.LayerDetails.RAUnit == AngleUnit.Hours)
                    {
                        // If hours data, send it as is
                        perspective.RightAscention = rowValueForRA;
                    }
                    else if (float.TryParse(rowValueForRA, out hours))
                    {
                        // If degrees data, divide it by 15 and send
                        perspective.RightAscention = (hours / 15).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        // Else send it as is as it cannot be converted to degrees
                        perspective.RightAscention = rowValueForRA;
                    }

                    perspective.Declination = rowData[layerMap.MappedColumnType.IndexOf(ColumnType.Dec)];
                }
                else
                {
                    if (referenceFramePath.StartsWith(Common.Constants.EarthFramePath, StringComparison.OrdinalIgnoreCase))
                    {
                        perspective = new Perspective(Common.Constants.EarthLookAt, referenceFrame, false, Common.Constants.LatitudeDefaultValue, Common.Constants.LongitudeDefaultValue, Common.Constants.ZoomDefaultValue, Common.Constants.RotationDefaultValue, Common.Constants.LookAngleDefaultValue, DateTime.Now.ToString(), Common.Constants.TimeRateDefaultValue, Common.Constants.EarthZoomTextDefaultValue, string.Empty);
                    }
                    else
                    {
                        perspective = new Perspective(Common.Constants.SolarSystemLookAt, referenceFrame, false, Common.Constants.LatitudeDefaultValue, Common.Constants.LongitudeDefaultValue, Common.Constants.ZoomDefaultValue, Common.Constants.RotationDefaultValue, Common.Constants.LookAngleDefaultValue, DateTime.Now.ToString(), Common.Constants.TimeRateDefaultValue, Common.Constants.EarthZoomTextDefaultValue, string.Empty);
                    }

                    perspective.Latitude = rowData[layerMap.MappedColumnType.IndexOf(ColumnType.Lat)];
                    perspective.Longitude = rowData[layerMap.MappedColumnType.IndexOf(ColumnType.Long)];
                }
            }

            // Update Perspective based on current state
            Perspective currentPerspective = null;
            try
            {
                currentPerspective = WWTManager.GetCameraView();
            }
            catch (CustomException)
            {
                // Ignore.
            }

            if (currentPerspective != null)
            {
                perspective.ObservingTime = currentPerspective.ObservingTime;
                perspective.TimeRate = currentPerspective.TimeRate;
                perspective.Rotation = currentPerspective.Rotation;
                perspective.LookAngle = currentPerspective.LookAngle;
                perspective.Zoom = currentPerspective.Zoom;
            }

            return perspective;
        }

        /// <summary>
        /// This function retrieves the last used and valid group.
        /// </summary>
        /// <param name="layerMap">Layer map object</param>
        /// <param name="isWWTRunning">Whether WWT is running or not?</param>
        /// <returns>
        /// A Group that represents the current last used group.
        /// </returns>
        private static Group GetLastUsedGroup(LayerMap layerMap, bool isWWTRunning)
        {
            ICollection<Group> groups = null;

            // Get the Groups from WWT only if WWT is running. Otherwise, initialize with an empty list.
            if (isWWTRunning)
            {
                groups = WWTManager.GetAllWWTGroups(true);
            }
            else
            {
                groups = new List<Group>();
            }

            Group lastUsed = null;
            if (layerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.Lat).Any()
                || layerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.Long).Any())
            {
                // if Lat/Lon are mapped 
                lastUsed = groups.GetDefaultEarthGroup();
            }
            else if (layerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.RA).Any()
                || layerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.Dec).Any())
            {
                // if RA/Dec are mapped
                lastUsed = groups.GetDefaultSkyGroup();
            }
            else
            {
                lastUsed = LastUsedGroup;
                if (lastUsed == null || lastUsed.IsDeleted || !WWTManager.IsValidGroup(lastUsed, groups))
                {
                    Perspective perspective = null;
                    try
                    {
                        // Get the perspective only if WWT is running.
                        if (isWWTRunning)
                        {
                            perspective = WWTManager.GetCameraView();
                        }
                    }
                    catch (CustomException)
                    {
                        // Ignore.
                    }

                    lastUsed = (perspective != null && !string.IsNullOrEmpty(perspective.ReferenceFrame)) ?
                        groups.SearchGroup(perspective.ReferenceFrame) : groups.GetDefaultEarthGroup();
                }
            }

            if (LastUsedGroup != lastUsed)
            {
                LastUsedGroup = lastUsed;
            }

            return lastUsed;
        }

        /// <summary>
        /// Shows validation error for Viewpoint from data functionality
        /// </summary>
        /// <param name="layerMap">layerMap instance</param>
        private static void ShowViewpointFromDataValidationError(LayerMap layerMap)
        {
            if (layerMap.IsXYZLayer())
            {
                Ribbon.ShowError(Resources.GotoViewpointfromXYZError);
            }
            else if (layerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.Geo).Any())
            {
                Ribbon.ShowError(Resources.GotoViewpointFromGeometryDataError);
            }
            else
            {
                if (layerMap.LayerDetails.Group.IsPlanet())
                {
                    //// Lat, Lon Error
                    Ribbon.ShowError(Resources.GotoViewpointFromDataForSunError);
                }
                else
                {
                    // RA, Dec Error
                    Ribbon.ShowError(Resources.GotoViewpointFromDataForSkyError);
                }
            }
        }

        /// <summary>
        /// Goes to viewpoint on view in WWT click
        /// </summary>
        /// <param name="selectedLayerMap">layer map object</param>
        private static void GotoViewpointOnViewInWWT(LayerMap selectedLayerMap)
        {
            try
            {
                // Check if Lat/Lon, RA/Dec are mapped or not
                if ((selectedLayerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.Lat).Any()
                        && selectedLayerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.Long).Any())
                    || (selectedLayerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.RA).Any()
                        && selectedLayerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.Dec).Any()))
                {
                    var rowData = selectedLayerMap.RangeName.RefersToRange.GetFirstDataRow();
                    if (rowData != null)
                    {
                        // Build perspective based on mappings
                        var perspective = GetPerspectiveFromLayerRowData(selectedLayerMap, rowData);
                        if (perspective != null)
                        {
                            GotoViewpointFromData(perspective);
                        }
                    }
                }
                else if (selectedLayerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.Geo).Any() || selectedLayerMap.IsXYZLayer())
                {
                    var lookAt = selectedLayerMap.GetLookAt();
                    WWTManager.SetMode(lookAt);
                }
            }
            catch (CustomException)
            {
                // Consume as error message should not be shown
            }
        }

        /// <summary>
        /// Get active worksheet from the workbook
        /// </summary>
        /// <returns>Empty worksheet</returns>
        private static _Worksheet GetActiveWorksheet()
        {
            _Worksheet workSheet = (_Worksheet)ThisAddIn.ExcelApplication.ActiveWorkbook.ActiveSheet;
            return workSheet;
        }

        /// <summary>
        /// Gets selected layer's worksheet for local in WWT layer
        /// </summary>
        /// <param name="selectedLayerMap">Selected layer map</param>
        /// <returns>Worksheet for the selected layer's range</returns>
        private static _Worksheet GetSelectedLayerWorksheet(LayerMap selectedLayerMap)
        {
            _Worksheet worksheet = null;

            // Validates if the selected layer map's range name is valid.
            if (selectedLayerMap.RangeName.IsValid() && selectedLayerMap.RangeName.RefersToRange.IsValid())
            {
                // Activate the sheet in which the range is present.
                if (selectedLayerMap.RangeName.RefersToRange.Worksheet != ThisAddIn.ExcelApplication.ActiveSheet)
                {
                    worksheet = ((_Worksheet)selectedLayerMap.RangeName.RefersToRange.Worksheet);
                }
                if (worksheet == null)
                {
                    worksheet = (_Worksheet)ThisAddIn.ExcelApplication.ActiveSheet;
                }

                worksheet.Activate();
            }
            return worksheet;
        }

        private static StringBuilder BuildQuery(StationDataModel model)
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("http://www.iris.edu/ws/station/query?");
            if (!string.IsNullOrWhiteSpace(model.Network))
            {
                sb.Append(@"&net=" + model.Network);
            }

            if (!string.IsNullOrWhiteSpace(model.Station))
            {
                sb.Append(@"&sta=" + model.Station);
            }

            if (!string.IsNullOrWhiteSpace(model.Location))
            {
                sb.Append(@"&loc=" + model.Location);
            }

            if (!string.IsNullOrWhiteSpace(model.Channel))
            {
                sb.Append(@"&cha=" + model.Channel);
            }

            if (!string.IsNullOrWhiteSpace(model.StartDate) && !string.IsNullOrWhiteSpace(model.EndDate))
            {
                sb.Append(@"&timewindow=" + model.StartDate + "," + model.EndDate);
            }

            if (!string.IsNullOrWhiteSpace(model.Level))
            {
                sb.Append(@"&level=" + model.Level);
            }

            sb.Replace("query?&", "query?");
            return sb;
        }

        private static string ParseDateTime(string input)
        {
            DateTime time;
            DateTime now = DateTime.Now;
            if (DateTime.TryParse(input, out time))
            {
                //1991-11-19T00:00:00
                if (time > now)
                {
                    return now.ToString(@"yyyy-MM-ddThh:mm:ss");
                }
            }

            return input;

        }

        private static StringBuilder BuildEarthquakeQuery(EarthquakeDataModel model)
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("http://www.iris.edu/servlet/eventserver/eventsXML.do?Catalog=ALL&Contributor=ALL&magtype=ALL&LatMax=86.20&LatMin=-79.50&LonMax=180.00&LonMin=-180.00");
            if (!string.IsNullOrWhiteSpace(model.StartDate))
            {
                sb.Append(@"&StartDate=" + model.StartDate.Replace("/", ""));
            }

            if (!string.IsNullOrWhiteSpace(model.EndDate))
            {
                sb.Append(@"&StopDate=" + model.EndDate.Replace("/", ""));
            }

            if (!string.IsNullOrWhiteSpace(model.MagnitudeMax))
            {
                sb.Append(@"&MagMax=" + model.MagnitudeMax);
            }

            if (!string.IsNullOrWhiteSpace(model.MagnitudeMin))
            {
                sb.Append(@"&MagMin=" + model.MagnitudeMin);
            }

            if (!string.IsNullOrWhiteSpace(model.DepthMax))
            {
                sb.Append(@"&DepthMax=" + model.DepthMax);
            }

            if (!string.IsNullOrWhiteSpace(model.DepthMin))
            {
                sb.Append(@"&DepthMin=" + model.DepthMin);
            }

            sb.Append(@"&priority=" + model.SelectedPriority.Key);
            sb.Append(@"&PointsMax=" + model.SelectedDisplayCount.Key);

            return sb;
        }

        private string GetDatafromFeed(string response)
        {
            response = response.Replace("Â¬â Igualeja", "Igualeja");
            XmlDocument xmlDoc = new XmlDocument();
            XmlNamespaceManager xmlNsMgr = new XmlNamespaceManager(xmlDoc.NameTable);
            xmlNsMgr.AddNamespace("we", "http://www.data.scec.org/xml/station/");

            xmlDoc.LoadXml(response);
            XmlNodeList elements = xmlDoc.DocumentElement.SelectNodes("//we:StationEpoch", xmlNsMgr);
            rowCount = elements.Count;
            StringBuilder sb = new StringBuilder();
            int firstRow = 1;

            XmlNode childElement = elements.Item(0);
            columnCount = childElement.ChildNodes.Count;

            tempValue = new object[rowCount, columnCount];
            //int row = 0;
            foreach (XmlNode element in elements)
            {
                XmlNodeList properties = element.ChildNodes;
                int count = 1;
                foreach (XmlNode property in properties)
                {
                    if (count != 1)
                    {
                        sb.Append("\t");
                    }

                    if (firstRow == 1)
                    {
                        string value = property.Name.Substring(property.Name.IndexOf(":") + 1, property.Name.Length - property.Name.IndexOf(":") - 1).Trim();
                        sb.Append(ParseDateTime(value));
                        //columnCount = properties.Count;  
                        tempValue[firstRow - 1, count - 1] = ParseDateTime(value);
                    }
                    else
                    {
                        string value = property.InnerText.Trim();
                        sb.Append(ParseDateTime(value));
                        tempValue[firstRow - 1, count - 1] = ParseDateTime(value);
                    }

                    count++;
                }

                firstRow++;
                //row++;
                sb.AppendLine(string.Empty);
            }

            return sb.ToString();
        }

        private string GetEarthquakeDatafromFeed(string response)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(response);
            XmlNodeList elements = xmlDoc.DocumentElement.SelectNodes("//event");
            rowCount = elements.Count;
            StringBuilder sb = new StringBuilder();
            int firstRow = 1;
            columnCount = elements[0].Attributes.Count;

            tempValue = new object[rowCount, columnCount];
            //int row = 0;
            foreach (XmlNode element in elements)
            {
                XmlAttributeCollection properties = element.Attributes;
                int count = 1;
                foreach (XmlAttribute property in properties)
                {
                    if (count != 1)
                    {
                        sb.Append("\t");
                    }

                    if (firstRow == 1)
                    {
                        string value = property.Name.Substring(property.Name.IndexOf(":") + 1, property.Name.Length - property.Name.IndexOf(":") - 1).Trim();
                        sb.Append(value);
                        //columnCount = properties.Count;  
                        tempValue[firstRow - 1, count - 1] = value;
                    }
                    else
                    {
                        string value = property.InnerText.Trim();
                        sb.Append(value);
                        tempValue[firstRow - 1, count - 1] = value;
                    }

                    count++;
                }

                firstRow++;
                //row++;
                sb.AppendLine(string.Empty);
            }

            return sb.ToString();
        }

        private bool HasThresholdCrossedWarningLimit()
        {
            const double WARNING_THRESHOLD_LIMIT = 20000;
            try
            {
                //Get magnitude Column num "from selected data" (Excel have 4 data columns but only last 3 being selected)
                int magColNum = -1;
                for (int currIndex = 0; currIndex < this.currentWorkbookMap.SelectedLayerMap.MappedColumnType.Count; currIndex++)
                {
                    if (this.currentWorkbookMap.SelectedLayerMap.MappedColumnType[currIndex] == ColumnType.Mag)
                    {
                        magColNum = currIndex + 1;
                        break;
                    }
                }
                //No Magnitude entered (What is default value for magnitude?)
                if (magColNum == -1)
                    return false;

                //get magnitude sum of the selected data
                Worksheet ws = ThisAddIn.ExcelApplication.ActiveSheet as Worksheet;
                Range selectedRange = ThisAddIn.ExcelApplication.Selection as Range;
                Range magnitudeColumnRange = selectedRange.Columns[magColNum, Type.Missing] as Range;

                Range firstCell = magnitudeColumnRange.Cells[1, 1] as Range;
                Range lastCell = magnitudeColumnRange.Cells[selectedRange.Rows.Count, 1] as Range;

                //Get last used row in the sheet
                int lastUsedRowNum = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1;

                //temporary cell to store sum of magnitude
                Range sumCell = ws.Cells[lastUsedRowNum + 1, 1] as Range;
                //set sum formula for the cell
                string formula = String.Concat("=sum(", firstCell.Address, ":", lastCell.Address, ")");
                sumCell.Formula = formula;

                string strMagSum = sumCell.Value2.ToString();
                double magSum = Convert.ToDouble(strMagSum);
                
                //clear formula
                sumCell.Formula = string.Empty;

                //calculate threshold
                double scaleFactor = this.currentWorkbookMap.SelectedLayerMap.LayerDetails.ScaleFactor;
                double threshold = scaleFactor * magSum;

                //compare threshold with warning threshold limit
                if (threshold > WARNING_THRESHOLD_LIMIT)
                    return true;
            }
            catch (Exception)
            {
            }

            return false;
        }

        /// <summary>
        /// This method will be called when the user click on View in WWT button.
        /// </summary>
        private void ViewInWWT()
        {
            if (this.currentWorkbookMap != null && this.currentWorkbookMap.SelectedLayerMap != null)
            {
                LayerMap selectedLayerMap = this.currentWorkbookMap.SelectedLayerMap;
                if (ValidateMappedColumns())
                {
                    try
                    {
                        //Check if Threshold has crossed warning limit.
                        if (this.layerDetailsViewModel.IsRenderingTimeoutAlertShown && HasThresholdCrossedWarningLimit())
                        {
                            string userMessage = Resources.RenderingThresholdWarning;
                            userMessage = userMessage.Replace("__NEW_LINE__", Environment.NewLine);

                            DialogResult result = MessageBox.Show(userMessage, "Warning!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                            if(result == DialogResult.No)
                                return;
                        }

                        Utility.IsWWTInstalled();
                        WWTManager.IsValidMachine(Common.Globals.TargetMachine.MachineIP.ToString(), false);
                        ThisAddIn.ExcelApplication.Cursor = XlMousePointer.xlWait;
                        this.SetCoordinateType();
                        if (CreateIfNotExist(selectedLayerMap))
                        {
                            if (UpdateWWT(selectedLayerMap))
                            {
                                selectedLayerMap.IsNotInSync = false;

                                // Set the view in WWT visibility and set selected layer name
                                this.SetLayerDetailsViewModelProperties();

                                SetGetLayerDataDisplayName(this.currentWorkbookMap.SelectedLayerMap);

                                // Save the details of the workbook map to the custom xml parts
                                ThisAddIn.ExcelApplication.ActiveWorkbook.SaveWorkbookMap(this.workBookMaps);

                                if (this.currentWorkbookMap.SelectedLayerMap.IsVisualizeClicked)
                                {
                                    this.currentWorkbookMap.SelectedLayerMap.IsVisualizeClicked = false;
                                }

                                // Show the layer manager pane in WWT.
                                WWTManager.ShowLayerManager();

                                // Activate the layer in WWT.
                                WWTManager.ActivateLayer(selectedLayerMap.LayerDetails.ID);

                                // We need to move to data in WWT only if "Auto Move" toggle button is enabled.
                                if (this.ribbonInstance.IsAutoMoveEnabled)
                                {
                                    GotoViewpointOnViewInWWT(selectedLayerMap);
                                }

                                if (TargetMachine.DefaultIP.ToString() == Common.Globals.TargetMachine.MachineIP.ToString())
                                {
                                    Utility.ShowWWT();
                                }
                            }
                            else
                            {
                                selectedLayerMap.IsNotInSync = true;
                            }
                        }
                    }
                    catch (OutOfMemoryException)
                    {
                        Ribbon.ShowError(Resources.DefaultErrorMessage);
                    }
                    catch (CustomException ex)
                    {
                        this.SyncAndRebindOnError(ex);
                    }
                    finally
                    {
                        ThisAddIn.ExcelApplication.Cursor = XlMousePointer.xlDefault;
                    }
                }
            }
        }

        private object[,] UpdatedWWTData(UpdateDataModel input)
        {
            object[,] actualData = null;
             
            object[,] data = currentWorksheet.UsedRange.GetDataArray(false);
            var header = currentWorksheet.UsedRange.GetHeader();
            if (data != null)
            {
                int rowCount = data.GetLength(0);
                int columnCount = data.GetLength(1);

                int latColumnIndex = 0;
                int longColumnIndex = 1;

                int colorColumnIndex = 2;
                int altitudeColumnIndex = 2;

                int rColumnIndex = 0;
                int gColumnIndex = 0;
                int bColumnIndex = 0;


                foreach (var item in header)
                {
                    if (Common.Constants.LatSearchList.Contains(item.ToLower()))
                    {
                        latColumnIndex = header.IndexOf(item);
                    }

                    if (Common.Constants.LonSearchList.Contains(item.ToLower()))
                    {
                        longColumnIndex = header.IndexOf(item);
                    }

                    if (!string.IsNullOrWhiteSpace(input.ColorColumn) && string.Compare(input.ColorColumn.ToLower(), item, true) == 0)
                    {
                        colorColumnIndex = header.IndexOf(item);
                    }

                    if (string.Compare(input.AltitudeColumn.ToLower(), item, true) == 0)
                    {
                        altitudeColumnIndex = header.IndexOf(item);
                    }

                    if (!string.IsNullOrWhiteSpace(input.RColumn) && string.Compare(input.RColumn.ToLower(), item, true) == 0)
                    {
                        rColumnIndex = header.IndexOf(item);
                    }
                    if (!string.IsNullOrWhiteSpace(input.GColumn) && string.Compare(input.GColumn.ToLower(), item, true) == 0)
                    {
                        gColumnIndex = header.IndexOf(item);
                    }
                    if (!string.IsNullOrWhiteSpace(input.BColumn) && string.Compare(input.BColumn.ToLower(), item, true) == 0)
                    {
                        bColumnIndex = header.IndexOf(item);
                    }
                }

                if (input.FilterBetweenBoundaries)
                {
                    actualData = FilterDataWithInBoundaries(input, data, rowCount, columnCount, latColumnIndex, longColumnIndex, altitudeColumnIndex, rColumnIndex, gColumnIndex, bColumnIndex);
                }
                else
                {
                    actualData = GetdataWithoutFilter(input, data, rowCount, columnCount, latColumnIndex, longColumnIndex, altitudeColumnIndex, rColumnIndex, gColumnIndex, bColumnIndex);
                }
            }

            return actualData;
        }

        private static object[,] FilterDataWithInBoundaries(UpdateDataModel input, object[,] data, int rowCount, int columnCount, int latColumnIndex, int longColumnIndex, int altitudeColumnIndex, int rColumnIndex, int gColumnIndex, int bColumnIndex)
        {
            object[,] actualData = null;

            object[,] updatedData = new object[rowCount, columnCount + 3];

            // Update Header values
            for (int columnnIndex = 0; columnnIndex < columnCount; columnnIndex++)
            {
                updatedData[0, columnnIndex] = data[1, columnnIndex + 1];
            }

            updatedData[0, columnCount] = "Geometry";
            updatedData[0, columnCount + 1] = "Color";
            updatedData[0, columnCount + 2] = "Altitude";

            int actualRowCount = 0;
            // Generate and update actual Data Values
            for (int rowIndex = 1; rowIndex < rowCount; rowIndex++)
            {
                if (UpdateDataModel.CheckWithinBoundary(
                    Convert.ToDouble(data[rowIndex + 1, latColumnIndex + 1].ToString()),
                    Convert.ToDouble(data[rowIndex + 1, longColumnIndex + 1].ToString()),
                    input))
                {
                    actualRowCount++;

                    for (int columnnIndex = 0; columnnIndex < columnCount; columnnIndex++)
                    {
                        updatedData[actualRowCount, columnnIndex] = data[rowIndex + 1, columnnIndex + 1];
                    }

                    // Update Geometry
                    updatedData[actualRowCount, columnCount] = UpdateDataModel.GetGeometry(
                        Convert.ToDouble(updatedData[actualRowCount, latColumnIndex].ToString()),
                        Convert.ToDouble(updatedData[actualRowCount, longColumnIndex].ToString()),
                        input);

                    // Update Color
                    updatedData[actualRowCount, columnCount + 1] = UpdateDataModel.GetColorValue(
                        Convert.ToInt16(updatedData[actualRowCount, rColumnIndex].ToString()),
                        Convert.ToInt16(updatedData[actualRowCount, gColumnIndex].ToString()),
                        Convert.ToInt16(updatedData[actualRowCount, bColumnIndex].ToString()));

                    // Update Altitude
                    updatedData[actualRowCount, columnCount + 2] = UpdateDataModel.GetAltitudeValue(
                        Convert.ToDouble(updatedData[actualRowCount, altitudeColumnIndex].ToString()),
                        input);
                }
            }

            actualData = new object[actualRowCount + 1, updatedData.GetLength(1)];

            for (int row = 0; row < actualRowCount; row++)
            {
                for (int columnn = 0; columnn < updatedData.GetLength(1); columnn++)
                {
                    actualData[row, columnn] = updatedData[row, columnn];
                }
            }
            return actualData;
        }

        private static object[,] GetdataWithoutFilter(UpdateDataModel input, object[,] data, int rowCount, int columnCount, int latColumnIndex, int longColumnIndex, int altitudeColumnIndex, int rColumnIndex, int gColumnIndex, int bColumnIndex)
        {
            object[,] updatedData = new object[rowCount, columnCount + 3];

            // Update Header values
            for (int columnnIndex = 0; columnnIndex < columnCount; columnnIndex++)
            {
                updatedData[0, columnnIndex] = data[1, columnnIndex + 1];
            }

            updatedData[0, columnCount] = "Geometry";
            updatedData[0, columnCount + 1] = "Color";
            updatedData[0, columnCount + 2] = "Altitude";

            // Generate and update actual Data Values
            for (int rowIndex = 1; rowIndex < rowCount; rowIndex++)
            {
                    for (int columnnIndex = 0; columnnIndex < columnCount; columnnIndex++)
                    {
                        updatedData[rowIndex, columnnIndex] = data[rowIndex + 1, columnnIndex + 1];
                    }

                    // Update Geometry
                    updatedData[rowIndex, columnCount] = UpdateDataModel.GetGeometry(
                        Convert.ToDouble(updatedData[rowIndex, latColumnIndex].ToString()),
                        Convert.ToDouble(updatedData[rowIndex, longColumnIndex].ToString()),
                        input);

                    // Update Color
                    updatedData[rowIndex, columnCount + 1] = UpdateDataModel.GetColorValue(
                        Convert.ToInt16(updatedData[rowIndex, rColumnIndex].ToString()),
                        Convert.ToInt16(updatedData[rowIndex, gColumnIndex].ToString()),
                        Convert.ToInt16(updatedData[rowIndex, bColumnIndex].ToString()));

                    // Update Altitude
                    updatedData[rowIndex, columnCount + 2] = UpdateDataModel.GetAltitudeValue(
                        Convert.ToDouble(updatedData[rowIndex, altitudeColumnIndex].ToString()),
                        input);
            }

            return updatedData;
        }

        private object[,] UpdatedHuricaneData(UpdateDataModel input, object[,] updatedData)
        {
            object[,] data = currentWorksheet.UsedRange.GetDataArray(false);
            var header = currentWorksheet.UsedRange.GetHeader();
            if (data != null)
            {
                int rowCount = data.GetLength(0);
                int columnCount = data.GetLength(1);

                int latColumnIndex = 0;
                int longColumnIndex = 1;
                int altitudeColumnIndex = 2;
                int sizeColumnIndex = 4;

                foreach (var item in header)
                {
                    if (Common.Constants.LatSearchList.Contains(item.ToLower()))
                    {
                        latColumnIndex = header.IndexOf(item);
                    }

                    if (Common.Constants.LonSearchList.Contains(item.ToLower()))
                    {
                        longColumnIndex = header.IndexOf(item);
                    }

                    if (string.Compare(input.AltitudeColumn.ToLower(), item, true) == 0)
                    {
                        altitudeColumnIndex = header.IndexOf(item);
                    }

                    if (string.Compare("Size (miles)", item, true) == 0)
                    {
                        sizeColumnIndex = header.IndexOf(item);
                    }
                }

                updatedData = new object[rowCount, columnCount + 3];

                // Update Header values
                for (int columnnIndex = 0; columnnIndex < columnCount; columnnIndex++)
                {
                    updatedData[0, columnnIndex] = data[1, columnnIndex + 1];
                }

                updatedData[0, columnCount] = "Geometry";
                updatedData[0, columnCount + 1] = "Altitude";

                // Generate and update actual Data Values
                for (int rowIndex = 1; rowIndex < rowCount; rowIndex++)
                {
                    for (int columnnIndex = 0; columnnIndex < columnCount; columnnIndex++)
                    {
                        updatedData[rowIndex, columnnIndex] = data[rowIndex + 1, columnnIndex + 1];
                    }

                    // Update Geometry
                    updatedData[rowIndex, columnCount] = UpdateDataModel.GetCircle(
                        Convert.ToDouble(updatedData[rowIndex, latColumnIndex].ToString()),
                        Convert.ToDouble(updatedData[rowIndex, longColumnIndex].ToString()),
                        Convert.ToDouble(updatedData[rowIndex, sizeColumnIndex].ToString()));

                    // Update Altitude
                    updatedData[rowIndex, columnCount + 1] = UpdateDataModel.GetAltitudeValue(
                        Convert.ToDouble(updatedData[rowIndex, altitudeColumnIndex].ToString()),
                        input);
                }
            }

            return updatedData;
        }

        /// <summary>
        /// Initializes the UpdateManager instance and downloadUpdatesViewModel.
        /// </summary>
        private void InitializeUpdateManager()
        {
            this.updateManager = new UpdateManager();
            this.downloadUpdatesViewModel = new DownloadUpdatesViewModel();

            // set the initial text for the download updates button on the task pane
            this.downloadUpdatesViewModel.IsDownloadUpdatesEnabled = false;
            this.downloadUpdatesViewModel.DownloadUpdatesLabel = Resources.DownloadUpdatesButtonLabel;
            this.downloadUpdatesViewModel.IsDownloadUpdatesVisible = false;
        }

        /// <summary>
        /// Sets the get layer data display name
        /// </summary>
        /// <param name="selectedLayerMap">Selected layer map</param>
        private void SetGetLayerDataDisplayName(LayerMap selectedLayerMap)
        {
            if (selectedLayerMap != null && this.layerDetailsViewModel != null)
            {
                switch (selectedLayerMap.MapType)
                {
                    case LayerMapType.Local:
                    case LayerMapType.WWT:
                        this.layerDetailsViewModel.LayerDataDisplayName = Resources.GetLayerData;
                        break;
                    case LayerMapType.LocalInWWT:
                        this.layerDetailsViewModel.LayerDataDisplayName = Resources.RefreshLayerData;
                        break;
                }
            }
        }

        /// <summary>
        /// Handle the custom exception when sync of layers and rebind of UI are required
        /// The error message is displayed in a message box.
        /// </summary>
        /// <param name="ex">The custom exception instance</param>
        private void SyncAndRebindOnError(CustomException ex)
        {
            Ribbon.ShowError(ex.HasCustomMessage ? ex.Message : Resources.DefaultErrorMessage);

            // Synchronize WWT layers.
            SyncOnWWTNotRunning();

            //// 2. Any change in Core OM needs to be saved
            ThisAddIn.ExcelApplication.ActiveWorkbook.SaveWorkbookMap(this.workBookMaps);

            // Rebind UI.
            this.BuildAndBindLayerDetailsViewModel();
        }

        /// <summary>
        /// This function is used to update the details of the object models based on the details which was updated.
        /// </summary>
        /// <param name="target">
        /// Updated range.
        /// </param>
        /// <param name="currentSheet">
        /// Current worksheet.
        /// </param>
        private void UpdateDetails(Range target, Worksheet currentSheet)
        {
            Dictionary<string, string> allNamedRange = new Dictionary<string, string>();
            bool buildAndBindLayerDetailsViewModel = false;

            // Get all Layer details into dictionary.
            this.currentWorkbookMap.LocalLayerMaps.ForEach(item =>
            {
                allNamedRange.Add(item.RangeDisplayName, item.RangeAddress);
            });

            // Get all affected ranges
            affectedNamedRanges = currentSheet.GetAffectedNamedRanges(target, allNamedRange);

            this.currentWorkbookMap.LocalLayerMaps.ForEach(
                item =>
                {
                    if (affectedNamedRanges.ContainsKey(item.RangeDisplayName) && item.RangeName.IsValid() && !handledNamedRanges.ContainsKey(item.RangeDisplayName))
                    {
                        handledNamedRanges.Add(item.RangeDisplayName, affectedNamedRanges[item.RangeDisplayName]);
                        buildAndBindLayerDetailsViewModel = true;
                        Range currentRange = currentSheet.Range[affectedNamedRanges[item.RangeDisplayName]];
                        Collection<string> latestHeader = currentRange.GetHeader();
                        Collection<string> previousHeader = item.HeaderRowData;

                        // We need to upload data first and then Update the header as to make sure that we update the latest header.
                        // If we update header and then push data which has header in it, 
                        //      there are chances that auto map in WWT might be called because we are including the header data.

                        // Update the data irrespective of header is changed or not.
                        UpdateData(currentRange, item);
                        if (item.RangeAddress != affectedNamedRanges[item.RangeDisplayName] && latestHeader.Count != previousHeader.Count)
                        {
                            // Header Change :- Scenarios in which this will occur.
                            // 1. When a Column is deleted.
                            // 2. When a Cells are deleted and shifted up/Down.
                            item.UpdateHeaderProperties(currentRange);
                        }
                        else if (latestHeader.Count != previousHeader.Count)
                        {
                            // Header Change :- Scenarios in which this will occur.
                            // 1. When a data in the columns header is updated 
                            item.UpdateHeaderProperties(currentRange);
                        }
                        else
                        {
                            for (int index = 0; index < latestHeader.Count; index++)
                            {
                                if (string.Compare(latestHeader[index], previousHeader[index], StringComparison.Ordinal) != 0)
                                {
                                    // Scenario in which this check is MUST: Headers are mapped with only data and not to header text and
                                    // any of the date time columns header row having date formatted as Text.
                                    DateTime latestDate = DateTime.MinValue, previousDate = DateTime.MinValue;

                                    // For Excel 2010, current range will return the date as Double value, which needs to be converted to 
                                    // date time first and then to be compared with previous date.
                                    double latestHeaderDate;
                                    if (Double.TryParse(latestHeader[index], out latestHeaderDate))
                                    {
                                        latestHeader[index] = DateTime.FromOADate(latestHeaderDate).ToString();
                                    }

                                    // In case of column format is set as only hh:mm:ss, then previous header date columns will
                                    // be returned as double value.
                                    double previousHeaderDate;
                                    if (Double.TryParse(previousHeader[index], out previousHeaderDate))
                                    {
                                        previousHeader[index] = DateTime.FromOADate(previousHeaderDate).ToString();
                                    }

                                    if (DateTime.TryParse(latestHeader[index], out latestDate) && DateTime.TryParse(previousHeader[index], out previousDate))
                                    {
                                        if (latestDate == previousDate)
                                        {
                                            continue;
                                        }
                                    }

                                    // Header Change :- Scenarios in which this will occur.
                                    // 1. When a data in the columns header is updated 
                                    item.UpdateHeaderProperties(currentRange);
                                    break;
                                }
                            }
                        }

                        // Sets co-ordinate type (Spherical/Rectangular) based on the mapping, If Lat/Long/RA/Dec is mapped,
                        // spherical takes preference over rectangular.
                        this.SetCoordinateType();

                        // Update the header details on every operation.
                        UpdateHeader(item);

                        // Update the actual Range address (This is required if row data is deleted/inserted)
                        item.RangeAddress = affectedNamedRanges[item.RangeDisplayName];
                    }
                });

            if (buildAndBindLayerDetailsViewModel)
            {
                // Rebind UI on if there are any affected ranges.
                this.BuildAndBindLayerDetailsViewModel(false);
            }
        }

        /// <summary>
        /// Saved Viewpoint map into Custom xml parts in workbook
        /// </summary>
        /// <param name="workbook">workbook instance</param>
        private void SaveViewpointMap(Workbook workbook)
        {
            var viewpointMap = this.viewpointMaps.Find(item => item.Workbook == workbook);
            if (viewpointMap != null)
            {
                if (viewpointMap.SerializablePerspective != null)
                {
                    string content = viewpointMap.Serialize();
                    if (!string.IsNullOrEmpty(content))
                    {
                        workbook.AddCustomXmlPart(content, Common.Constants.ViewpointMapXmlNamespace);
                    }
                }
            }
        }

        /// <summary>
        /// Syncs WorkbookMap collection with Excel workbook collection in case of 
        /// Open scenarios where existing excel workbook objects get overwritten.
        /// For example, Book1 gets overwritten by workbook which is getting opened
        /// </summary>
        private void SyncWorkbookMapWithExcel()
        {
            for (int count = this.workBookMaps.Count - 1; count >= 0; count--)
            {
                var workbookMap = this.workBookMaps[count];

                // Check if workbook map has workbook instance which still exists in excel workbook list
                bool workbookFound = false;
                foreach (Workbook workbook in ThisAddIn.ExcelApplication.Workbooks)
                {
                    if (workbookMap.Workbook == workbook)
                    {
                        workbookFound = true;
                        break;
                    }
                }

                // If not found, the workbook instance is removed/overwritten in excel
                // so remove it from workbook map
                if (!workbookFound)
                {
                    workbookMap.StopAllNotifications();
                    this.workBookMaps.Remove(workbookMap);
                }
            }
        }

        /// <summary>
        /// Syncs Viewpoint collection with Excel workbook collection in case of 
        /// Open scenarios where existing excel workbook objects get overwritten.
        /// For example, Book1 gets overwritten by workbook which is getting opened
        /// </summary>
        private void SyncViewpointMapWithExcel()
        {
            for (int count = this.viewpointMaps.Count - 1; count >= 0; count--)
            {
                var viewpointMap = this.viewpointMaps[count];

                // Check if workbook map has workbook instance which still exists in excel workbook list
                bool workbookFound = false;
                foreach (Workbook workbook in ThisAddIn.ExcelApplication.Workbooks)
                {
                    if (viewpointMap.Workbook == workbook)
                    {
                        workbookFound = true;
                        break;
                    }
                }

                // If not found, the workbook instance is removed/overwritten in excel
                // so remove it from workbook map
                if (!workbookFound)
                {
                    this.viewpointMaps.Remove(viewpointMap);
                }
            }
        }

        /// <summary>
        /// Attach event handlers for workbook events
        /// </summary>
        private void AttachWorkbookEventHandlers()
        {
            if (ThisAddIn.ExcelApplication != null)
            {
                ThisAddIn.ExcelApplication.WindowDeactivate += new AppEvents_WindowDeactivateEventHandler(OnWindowDeactivate);
                ThisAddIn.ExcelApplication.WorkbookOpen += new AppEvents_WorkbookOpenEventHandler(OnWorkbookOpen);
                ThisAddIn.ExcelApplication.WorkbookActivate += new AppEvents_WorkbookActivateEventHandler(OnWorkbookActivate);
                ThisAddIn.ExcelApplication.ProtectedViewWindowActivate += new AppEvents_ProtectedViewWindowActivateEventHandler(OnProtectedWorkbookActivate);

                ((Microsoft.Office.Interop.Excel.AppEvents_Event)ThisAddIn.ExcelApplication).NewWorkbook += new AppEvents_NewWorkbookEventHandler(OnNewWorkbook);
                ThisAddIn.ExcelApplication.SheetChange += new AppEvents_SheetChangeEventHandler(OnSheetChange);

                ThisAddIn.ExcelApplication.SheetActivate += new AppEvents_SheetActivateEventHandler(OnSheetActivate);
                ThisAddIn.ExcelApplication.SheetDeactivate += new AppEvents_SheetDeactivateEventHandler(OnSheetDeactivate);

                ThisAddIn.ExcelApplication.AfterCalculate += new AppEvents_AfterCalculateEventHandler(OnAfterCalculate);
            }
        }

        /// <summary>
        /// Attaches custom changed event to the custom task pane
        /// </summary>
        private void AttachCustomTaskEventHandlers()
        {
            this.layerDetailsViewModel.CustomTaskPaneStateChangedEvent += new EventHandler(OnCustomTaskPaneChangedState);
            this.layerDetailsViewModel.LayerSelectionChangedEvent += new EventHandler(OnLayerSelectionChanged);
            this.layerDetailsViewModel.ViewnInWWTClickedEvent += new EventHandler(OnViewInWWTClicked);
            this.layerDetailsViewModel.ShowRangeClickedEvent += new EventHandler(OnShowRangeClickedEvent);
            this.layerDetailsViewModel.DeleteMappingClickedEvent += new EventHandler(OnDeleteMappingClickedEvent);
            this.layerDetailsViewModel.GetLayerDataClickedEvent += new EventHandler(OnGetLayerDataClickedEvent);
            this.layerDetailsViewModel.RefreshDropDownClickedEvent += new EventHandler(OnRefreshDropDownClickedEvent);
            this.layerDetailsViewModel.UpdateLayerClickedEvent += new EventHandler(OnUpdateLayerClickedEvent);
            this.layerDetailsViewModel.RefreshGroupDropDownClickedEvent += new EventHandler(OnRefreshGroupDropDownClickedEvent);
            this.layerDetailsViewModel.ReferenceSelectionChanged += new EventHandler(OnReferenceSelectionChanged);
            this.layerDetailsViewModel.DownloadUpdatesClickedEvent += new EventHandler(OnDownloadUpdatesClicked);
        }

        /// <summary>
        /// Attaches event handlers for ribbon events
        /// </summary>
        private void AttachRibbonEventHandlers()
        {
            this.ribbonInstance.VisualizeSelectionClicked += new EventHandler(OnVisualizeSelectionClicked);
            this.ribbonInstance.TargetMachineChanged += new EventHandler(OnTargetMachineChanged);
            this.ribbonInstance.GetViewpointClicked += new EventHandler(OnGetViewpointClicked);
            this.ribbonInstance.GotoViewpointClicked += new EventHandler(OnGotoViewpointClicked);
            this.ribbonInstance.GotoViewpointFromDataClicked += new EventHandler(OnGotoViewpointFromDataClicked);
            this.ribbonInstance.ManageViewpointClicked += new EventHandler(OnManageViewpointClicked);
            this.ribbonInstance.DownloadUpdatesButtonClicked += new EventHandler(OnDownloadUpdatesClicked);
        }

        /// <summary>
        /// Attach event handlers for the update manager
        /// </summary>
        private void AttachUpdateManagerEventHandlers()
        {
            this.updateManager.UpdateAvailable += new EventHandler(this.OnUpdateAvailable);
            this.updateManager.DownloadCompleted += new EventHandler(this.OnDownloadCompleted);
            this.updateManager.InstallationCompleted += new EventHandler(this.OnInstallationCompleted);
        }

        /// <summary>
        /// Set layer details properties for the dropdown in layer details view model
        /// </summary>
        private void SetLayerDetailsViewModelProperties()
        {
            // Set the view in WWT visibility and set selected layer name
            layerDetailsViewModel.IsViewInWWTEnabled = false;
            layerDetailsViewModel.IsLayerInSyncInfoVisible = true;

            // Start the animation for the layer in sync text.
            layerDetailsViewModel.StartShowHighlightAnimationTimer();
            layerDetailsViewModel.IsCallOutVisible = false;
            layerDetailsViewModel.IsGetLayerDataEnabled = true;
            layerDetailsViewModel.IsReferenceGroupEnabled = this.currentWorkbookMap.SelectedLayerMap.IsLayerCreated();

            // Reset the layer dropdown properties for the view model
            layerDetailsViewModel.SetSelectedLayerValues(layerDetailsViewModel.SelectedLayerName);
        }

        /// <summary>
        /// Update custom task pane properties in WWT
        /// </summary>
        private void UpdateWWTPropertiesForSelectedLayer()
        {
            if (this.currentWorkbookMap != null && this.currentWorkbookMap.SelectedLayerMap != null)
            {
                LayerMap selectedLayerMap = this.currentWorkbookMap.SelectedLayerMap;
                if (selectedLayerMap.CanUpdateWWT())
                {
                    try
                    {
                        WWTManager.IsValidMachine(Common.Globals.TargetMachine.MachineIP.ToString(), false);

                        // This is used only for Layer Property update notifications. Since the notifications are handled in background thread,
                        // CodeUpdate cannot be set the false here (main thread). It needs to be reset only by background thread once the 
                        // notification is handled by the background thread.
                        selectedLayerMap.IsPropertyChangedFromCode = true;

                        // WWT updates to happen for non local layers
                        // Update the header details.
                        if (!WWTManager.UpdateLayer(selectedLayerMap.LayerDetails, true, false))
                        {
                            if (selectedLayerMap.MapType == LayerMapType.WWT)
                            {
                                // 1. Show Warning
                                Ribbon.ShowWarning(Properties.Resources.OnlyWWTLayerGotDeleted);

                                // 2. Remove the WWT layer and set the selected layer map to null.
                                this.currentWorkbookMap.AllLayerMaps.Remove(selectedLayerMap);
                                this.currentWorkbookMap.SelectedLayerMap = null;
                            }
                            else
                            {
                                // 1. Show Warning
                                Ribbon.ShowWarning(Properties.Resources.WWTLayerGotDeleted);

                                // 2. Set the IsNotInSync flag to true when you have data in Excel.
                                selectedLayerMap.IsNotInSync = true;
                            }

                            // Rebind UI
                            this.BuildAndBindLayerDetailsViewModel();
                        }
                        else
                        {
                            // If layer properties are updated correctly, activate the layer
                            WWTManager.ActivateLayer(selectedLayerMap.LayerDetails.ID);
                        }
                    }
                    catch (CustomException ex)
                    {
                        this.SyncAndRebindOnError(ex);
                    }
                }
            }
        }

        /// <summary>
        /// This function is used to select the range in Excel.
        /// </summary>
        private void ShowSelectedRange()
        {
            if (this.currentWorkbookMap != null && this.currentWorkbookMap.SelectedLayerMap != null)
            {
                LayerMap selectedLayerMap = this.currentWorkbookMap.SelectedLayerMap;
                if (IsLocalLayer(selectedLayerMap.MapType) && selectedLayerMap.RangeName.IsValid() && selectedLayerMap.RangeName.RefersToRange.IsValid())
                {
                    // Activate the sheet in which the range is present.
                    if (selectedLayerMap.RangeName.RefersToRange.Worksheet != ThisAddIn.ExcelApplication.ActiveSheet)
                    {
                        ((_Worksheet)selectedLayerMap.RangeName.RefersToRange.Worksheet).Activate();
                    }

                    selectedLayerMap.RangeName.RefersToRange.Select();
                }
            }
        }

        /// <summary>
        /// Deletes the mapping for local or local in WWT layers
        /// </summary>
        private void DeleteMapping()
        {
            if (this.currentWorkbookMap != null && this.currentWorkbookMap.SelectedLayerMap != null)
            {
                LayerMap selectedLayerMap = this.currentWorkbookMap.SelectedLayerMap;
                if (IsLocalLayer(selectedLayerMap.MapType))
                {
                    try
                    {
                        // If the layer is local in WWT and in synch, delete the layer from WWT. 
                        // Even if Layer is LocalInWWT and in sync, layer might have deleted in WWT after selecting the layer in layer dropdown.
                        if (selectedLayerMap.MapType == LayerMapType.LocalInWWT && !selectedLayerMap.IsNotInSync && WWTManager.IsValidLayer(selectedLayerMap.LayerDetails.ID))
                        {
                            WWTManager.DeleteLayer(selectedLayerMap.LayerDetails.ID);
                        }

                        // Range name is deleted
                        this.currentWorkbookMap.SelectedLayerMap.RangeName.Delete();

                        this.currentWorkbookMap.AllLayerMaps.Remove(selectedLayerMap);
                        this.currentWorkbookMap.SelectedLayerMap = null;
                        ThisAddIn.ExcelApplication.ActiveWorkbook.SaveWorkbookMap(this.workBookMaps);

                        // No need to build the reference frame dropdown while deleting the mapping.
                        this.BuildAndBindLayerDetailsViewModel(false);
                    }
                    catch (COMException ex)
                    {
                        Logger.LogException(ex);
                    }
                    catch (CustomException ex)
                    {
                        SyncAndRebindOnError(ex);
                    }
                }
            }
        }

        /// <summary>
        /// Get layer data for local in WWT and WWT layers
        /// </summary>
        private void GetLayerData()
        {
            if (this.currentWorkbookMap != null && this.currentWorkbookMap.SelectedLayerMap != null && this.currentWorkbookMap.SelectedLayerMap.LayerDetails != null
                && !string.IsNullOrEmpty(this.currentWorkbookMap.SelectedLayerMap.LayerDetails.ID))
            {
                if (WWTManager.IsValidLayer(this.currentWorkbookMap.SelectedLayerMap.LayerDetails.ID))
                {
                    try
                    {
                        _Worksheet workSheet = GetActiveWorksheet();

                        if (workSheet != null)
                        {
                            switch (this.currentWorkbookMap.SelectedLayerMap.MapType)
                            {
                                case LayerMapType.WWT:
                                    GetLayerDataForWWT(workSheet);
                                    break;
                                case LayerMapType.LocalInWWT:
                                    GetLayerDataForLocalInWWT(GetSelectedLayerWorksheet(this.currentWorkbookMap.SelectedLayerMap));
                                    break;
                            }
                        }
                    }
                    catch (CustomException ex)
                    {
                        Ribbon.ShowError(ex.HasCustomMessage ? ex.Message : Resources.DefaultErrorMessage);
                    }
                }
                else
                {
                    // Warning message for the layer getting deleted
                    Ribbon.ShowWarning(Properties.Resources.OnlyWWTLayerGotDeleted);

                    this.currentWorkbookMap.RefreshLayers();
                }

                // Any change in Core OM needs to be saved
                ThisAddIn.ExcelApplication.ActiveWorkbook.SaveWorkbookMap(this.workBookMaps);
                this.BuildAndBindLayerDetailsViewModel();
            }
        }

        /// <summary>
        /// Gets layer data for WWT layer
        /// </summary>
        /// <param name="workSheet">Active worksheet</param>
        private void GetLayerDataForWWT(_Worksheet workSheet)
        {
            if (workSheet != null)
            {
                object[,] layerData = WWTManager.GetLayerData(this.currentWorkbookMap.SelectedLayerMap.LayerDetails.ID, false);
                if (layerData != null && layerData.Length > 0)
                {
                    // Gets the range from the excel for data row and columns
                    Range currentRange = workSheet.GetRange(ThisAddIn.ExcelApplication.ActiveCell, layerData.GetLength(0), layerData.GetLength(1));
                    if (currentRange != null)
                    {
                        if (ValidateAffectedMappedRange(workSheet, currentRange))
                        {
                            string address = currentRange.Address;

                            if (currentRange != null)
                            {
                                currentRange.Select();
                                InsertRows(currentRange);

                                // Gets the new range for with the active cell address
                                Range newRange = workSheet.Application.Range[address];

                                // Creates named range for the new range
                                CreateRangeForLayer(newRange);
                                newRange.SetValue(layerData);
                                SetFormatForDateColumns(workSheet);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Sets the Date format for the columns which are mapped to StartDate and EndDate.
        /// </summary>
        /// <param name="workSheet">Worksheet object</param>
        private void SetFormatForDateColumns(_Worksheet workSheet)
        {
            Range currentLayerRange = workSheet.Application.Range[this.currentWorkbookMap.SelectedLayerMap.RangeAddress];
            Range firstCell = null;

            if (this.currentWorkbookMap.SelectedLayerMap.LayerDetails.StartDateColumn != Common.Constants.DefaultColumnIndex)
            {
                foreach (Range area in currentLayerRange.Areas)
                {
                    firstCell = area.Cells[this.currentWorkbookMap.SelectedLayerMap.LayerDetails.StartDateColumn + 1] as Range;
                    Range startDateColumn = workSheet.GetRange(firstCell, currentLayerRange.GetRowsCount(), 1);
                    startDateColumn.NumberFormat = "m/d/yyyy h:mm";
                }
            }

            if (this.currentWorkbookMap.SelectedLayerMap.LayerDetails.EndDateColumn != Common.Constants.DefaultColumnIndex)
            {
                foreach (Range area in currentLayerRange.Areas)
                {
                    firstCell = area.Cells[this.currentWorkbookMap.SelectedLayerMap.LayerDetails.EndDateColumn + 1] as Range;
                    Range endDateColumn = workSheet.GetRange(firstCell, currentLayerRange.GetRowsCount(), 1);
                    endDateColumn.NumberFormat = "m/d/yyyy h:mm";
                }
            }
        }

        /// <summary>
        /// Gets the layer data for local in WWT
        /// </summary>
        /// <param name="workSheet">Current active worksheet</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1814:PreferJaggedArraysOverMultidimensional", MessageId = "Body", Justification = "We cannot use jagged array in this scenario because the excel Object model is designed to convert the value as [,].")]
        private void GetLayerDataForLocalInWWT(_Worksheet workSheet)
        {
            if (workSheet != null && this.currentWorkbookMap.SelectedLayerMap != null)
            {
                Range currentLayerRange = workSheet.Application.Range[this.currentWorkbookMap.SelectedLayerMap.RangeAddress];
                Range validationRange = currentLayerRange;
                if (currentLayerRange != null)
                {
                    object[,] layerData = WWTManager.GetLayerData(this.currentWorkbookMap.SelectedLayerMap.LayerDetails.ID, false);

                    if (layerData != null && layerData.Length > 0)
                    {
                        int difference = 0;
                        Range rowDifferenceRange = null, columnDifferenceRange = null;
                        string rowaddress = string.Empty, colAddress = string.Empty;

                        // 1. Gets the difference from the current layer range 
                        // 2. Gets the range with the first  row cell
                        // 3. Inserts rows/columns to range with difference
                        if (currentLayerRange.Columns.Count < layerData.GetLength(1))
                        {
                            difference = GetRangeDifference(currentLayerRange.Columns.Count, layerData.GetLength(1));

                            // Take the column next to last column of the current range as first column getting inserted.
                            Range firstColumnCell = (Range)currentLayerRange.Cells[1, currentLayerRange.Columns.Count + 1];

                            columnDifferenceRange = workSheet.GetRange(firstColumnCell, currentLayerRange.GetRowsCount(), difference);
                        }

                        // The range provides a combination of the current range and column difference range,
                        // which is used for affected range and formula validation.
                        if (difference > 0)
                        {
                            validationRange = validationRange.Resize[Type.Missing, validationRange.Columns.Count + difference];
                        }

                        int rowsCount = currentLayerRange.GetRowsCount();
                        if (rowsCount < layerData.GetLength(0))
                        {
                            difference = GetRangeDifference(rowsCount, layerData.GetLength(0));

                            // Take the row next to last row of the current range as first row getting inserted.
                            Range firstRowCell = (Range)currentLayerRange.Cells[rowsCount + 1, 1];

                            rowDifferenceRange = workSheet.GetRange(firstRowCell, difference, currentLayerRange.Columns.Count);
                        }

                        // The range provides a combination of the current range and row difference range
                        // which is used for affected range and formula validation.
                        if (difference > 0)
                        {
                            validationRange = validationRange.Resize[validationRange.GetRowsCount() + difference, Type.Missing];
                        }

                        if (ValidateLocalInWWTLayerData(workSheet, validationRange))
                        {
                            Range rowRange = null, colRange = null;
                            if (columnDifferenceRange != null)
                            {
                                // Insert the columns which are needed to get the latest data.
                                colAddress = columnDifferenceRange.Address;
                                InsertColumns(columnDifferenceRange);
                                colRange = workSheet.Application.Range[colAddress];
                            }

                            if (rowDifferenceRange != null)
                            {
                                // Insert the rows which are needed to get the latest data.
                                rowaddress = rowDifferenceRange.Address;
                                InsertRows(rowDifferenceRange);
                                rowRange = workSheet.Application.Range[rowaddress];
                            }

                            bool createLayerForRange = false;
                            if (rowRange != null)
                            {
                                // Any new rows added, add them to the layer range.
                                createLayerForRange = true;
                                currentLayerRange = workSheet.Application.get_Range(rowRange, currentLayerRange);
                            }

                            if (colRange != null)
                            {
                                // Any new columns added, add them to the layer range.
                                createLayerForRange = true;
                                currentLayerRange = workSheet.Application.get_Range(colRange, currentLayerRange);
                            }
                            ThisAddIn.ExcelApplication.SheetChange -= new AppEvents_SheetChangeEventHandler(OnSheetChange);
                            currentLayerRange.Cells.Clear();
                            ThisAddIn.ExcelApplication.SheetChange += new AppEvents_SheetChangeEventHandler(OnSheetChange);
                            currentLayerRange.Select();

                            // Only in case if the range size increased, update the layer properties.
                            if (createLayerForRange)
                            {
                                CreateRangeForLayer(currentLayerRange);
                            }
                            else
                            {
                                // Since Range is not getting changed in case of no more rows/columns added to them, call this method to update the layer properties alone.
                                SetLayerRangeProperties(this.currentWorkbookMap.SelectedLayerMap.RangeName);
                            }

                            bool isResizeRangeRequired = false;

                            // In case if the rows are less, resize the range, instead paste only the rows having data and rest of rows will be left as empty.
                            // Additional empty rows cannot be deleted, since there could be data in the columns which are not part of range which will be lost.
                            if (currentLayerRange.GetRowsCount() > layerData.GetLength(0))
                            {
                                difference = GetRangeDifference(currentLayerRange.GetRowsCount(), layerData.GetLength(0));
                                currentLayerRange = currentLayerRange.Resize[currentLayerRange.GetRowsCount() - difference, Type.Missing];
                                isResizeRangeRequired = true;
                            }

                            // In case if the columns are less, resize the range, instead paste only the columns having data and rest of columns will be left as empty.
                            // Additional empty columns cannot be deleted, since there could be data in the rows which are not part of range which will be lost.
                            if (currentLayerRange.Columns.Count > layerData.GetLength(1))
                            {
                                difference = GetRangeDifference(currentLayerRange.Columns.Count, layerData.GetLength(1));
                                currentLayerRange = currentLayerRange.Resize[Type.Missing, currentLayerRange.Columns.Count - difference];
                                isResizeRangeRequired = true;
                            }

                            if (isResizeRangeRequired)
                            {
                                CreateRangeForLayer(currentLayerRange);
                                currentLayerRange.Select();
                            }

                            currentLayerRange.SetValue(layerData);
                            SetFormatForDateColumns(workSheet);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Validates if the range is affecting any of the existing range of layers
        /// and checks if the range has formula
        /// </summary>
        /// <param name="workSheet">Current active worksheet</param>
        /// <param name="range">Current range</param>
        /// <returns>Returns true if the layer is valid</returns>
        private bool ValidateLocalInWWTLayerData(_Worksheet workSheet, Range range)
        {
            bool isValid = false;
            if (ValidateAffectedMappedRange(workSheet, range))
            {
                if (range.ValidateFormula())
                {
                    isValid = Ribbon.ShowWarningWithResult(Properties.Resources.RangeHasFormulaWarning);
                }
                else
                {
                    isValid = true;
                }
            }
            return isValid;
        }

        /// <summary>
        /// Creates a named range for the selected layer and updates the 
        /// selection range.
        /// </summary>
        /// <param name="range">Updated selection range</param>
        private void CreateRangeForLayer(Range range)
        {
            string selectionRangeName = ThisAddIn.ExcelApplication.ActiveWorkbook.GetSelectionRangeName();
            if (!string.IsNullOrEmpty(selectionRangeName))
            {
                Name namedRange = ThisAddIn.ExcelApplication.ActiveWorkbook.CreateNamedRange(selectionRangeName, range);
                if (namedRange != null)
                {
                    SetLayerRangeProperties(namedRange);
                }
            }
        }

        /// <summary>
        /// Inserts range to specified range
        /// </summary>
        /// <param name="range">Range to insert row</param>
        /// <returns>Range with inserted row</returns>
        private Range InsertRows(Range range)
        {
            if (range != null)
            {
                // Remove sheet change event from excel so that data is not updated in WWT
                ThisAddIn.ExcelApplication.SheetChange -= new AppEvents_SheetChangeEventHandler(OnSheetChange);

                // Clears the content which was copied from the same workbook, so that data is not used from copied content.               
                Clipboard.Clear();
                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                ThisAddIn.ExcelApplication.SheetChange += new AppEvents_SheetChangeEventHandler(OnSheetChange);
            }

            return range;
        }

        /// <summary>
        /// Inserts range to specified range
        /// </summary>
        /// <param name="range">Range to insert column</param>
        /// <returns>Range with inserted column</returns>
        private Range InsertColumns(Range range)
        {
            if (range != null)
            {
                // Remove sheet change event from excel so that data is not updated in WWT
                ThisAddIn.ExcelApplication.SheetChange -= new AppEvents_SheetChangeEventHandler(OnSheetChange);
                range.EntireColumn.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                ThisAddIn.ExcelApplication.SheetChange += new AppEvents_SheetChangeEventHandler(OnSheetChange);
            }

            return range;
        }

        /// <summary>
        /// Validates the affected mapped range to check if the data is going to be imported
        /// in any of the available range
        /// </summary>
        /// <param name="currentSheet">Current active sheet</param>
        /// <param name="currentRange">Target range</param>
        /// <returns>True if the data is in affected range</returns>
        private bool ValidateAffectedMappedRange(_Worksheet currentSheet, Range currentRange)
        {
            bool isContinue = true;
            if (currentSheet != null && currentRange != null && this.currentWorkbookMap != null
                && this.currentWorkbookMap.SelectedLayerMap != null)
            {
                Dictionary<string, string> allNamedRange = new Dictionary<string, string>();

                // Get all Layer details into dictionary.
                this.currentWorkbookMap.LocalLayerMaps.ForEach(item =>
                {
                    allNamedRange.Add(item.RangeDisplayName, item.RangeAddress);
                });
                Dictionary<string, string> affectedMappedNamedRanges = currentSheet.GetAffectedNamedRanges(currentRange, allNamedRange);
                foreach (LayerMap layerMap in this.currentWorkbookMap.LocalLayerMaps)
                {
                    if (this.currentWorkbookMap.SelectedLayerMap.MapType == LayerMapType.WWT)
                    {
                        // Checks if the affected range is in any of the existing ranges for WWT layer
                        if (affectedMappedNamedRanges.ContainsKey(layerMap.RangeDisplayName) && layerMap.RangeName.IsValid())
                        {
                            Ribbon.ShowError(Properties.Resources.MappedWWTRangeError);
                            isContinue = false;
                            break;
                        }
                    }
                    else
                    {
                        // Checks if the affected range is in any of the existing ranges for local in WWT layer and 
                        // the range is not the current range for selected layer
                        if (affectedMappedNamedRanges.ContainsKey(layerMap.RangeDisplayName) && layerMap.RangeName.IsValid() && !layerMap.RangeDisplayName.Equals(this.currentWorkbookMap.SelectedLayerMap.RangeDisplayName))
                        {
                            isContinue = Ribbon.ShowWarningWithResult(Properties.Resources.MappedLocalInWWTRangeError);
                            break;
                        }
                    }
                }
            }
            else
            {
                isContinue = false;
            }
            return isContinue;
        }

        /// <summary>
        /// Sets range properties for the selected layer
        /// </summary>
        /// <param name="namedRange">Named range</param>
        private void SetLayerRangeProperties(Name namedRange)
        {
            if (this.currentWorkbookMap != null && this.currentWorkbookMap.SelectedLayerMap != null)
            {
                // Gets the selected layer
                LayerMap selectedLayerMap = this.currentWorkbookMap.AllLayerMaps.Where(layerMap => layerMap.LayerDetails.ID == this.currentWorkbookMap.SelectedLayerMap.LayerDetails.ID).FirstOrDefault();
                if (selectedLayerMap != null)
                {
                    // Resets the range name with the new name for the selected layer.
                    selectedLayerMap.ResetRange(namedRange);
                    if (selectedLayerMap.MapType == LayerMapType.WWT)
                    {
                        selectedLayerMap.MapType = LayerMapType.LocalInWWT;
                    }
                    selectedLayerMap.IsNotInSync = false;
                    this.currentWorkbookMap.SelectedLayerMap = selectedLayerMap;

                    Layer layer = WWTManager.GetLayerDetails(this.currentWorkbookMap.SelectedLayerMap.LayerDetails.ID, selectedLayerMap.LayerDetails.Group, false);
                    if (layer != null)
                    {
                        // Set other properties required for WWT map type
                        this.currentWorkbookMap.SelectedLayerMap = this.currentWorkbookMap.SelectedLayerMap.UpdateLayerMapProperties(layer);
                    }
                }
            }
        }

        /// <summary>
        /// This function is used to do Clean Up of WWT Layers when WWT is not running.
        /// </summary>
        private void SyncOnWWTNotRunning()
        {
            // 1. Cleanup all WWT layers.
            this.currentWorkbookMap.CleanUpWWTLayers();
        }

        /// <summary>
        /// Validates the mapped columns for mandatory mappings
        /// Either latitude and longitude column has to be mapped or 
        /// geometry column has to be mapped.
        /// </summary>
        /// <returns>If the mandatory columns are mapped.</returns>
        private bool ValidateMappedColumns()
        {
            bool isMandatoryColMapped = true;

            if (!((this.currentWorkbookMap.SelectedLayerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.Lat).Any()
                    && this.currentWorkbookMap.SelectedLayerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.Long).Any())
                || (this.currentWorkbookMap.SelectedLayerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.RA).Any()
                    && this.currentWorkbookMap.SelectedLayerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.Dec).Any())
                || this.currentWorkbookMap.SelectedLayerMap.IsXYMappedLayer()
                    || (this.currentWorkbookMap.SelectedLayerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.Geo).Any())))
            {
                string errorMessage = string.Empty;
                if (this.currentWorkbookMap.SelectedLayerMap.IsXYZLayer())
                {
                    errorMessage = Properties.Resources.MappedXYZColumnValidationError;
                }
                else
                {
                    errorMessage = this.currentWorkbookMap.SelectedLayerMap.LayerDetails.Group.IsPlanet() ?
                            Properties.Resources.MappedSunColumnValidationError : Properties.Resources.MappedSkyColumnValidationError;
                }

                isMandatoryColMapped = Ribbon.ShowWarningWithResult(errorMessage);
            }

            return isMandatoryColMapped;
        }

        /// <summary>
        /// Sets co-ordinate type(Rectangular/Spherical) based on the mappings
        /// If x,y,z is mapped and no Lat/Long or RA/DEC is mapped then co-ordinate type is Rectangular
        /// else co-ordinate type is Spherical
        /// </summary>
        private void SetCoordinateType()
        {
            if (this.currentWorkbookMap.SelectedLayerMap.IsXYZLayer())
            {
                this.currentWorkbookMap.SelectedLayerMap.LayerDetails.CoordinatesType = CoordinatesType.Rectangular;
            }
            else
            {
                this.currentWorkbookMap.SelectedLayerMap.LayerDetails.CoordinatesType = CoordinatesType.Spherical;
            }
        }

        /// <summary>
        /// This function is called when the user clicks on the layer maps dropdown.
        /// </summary>
        private void OnRefreshMapping()
        {
            try
            {
                ThisAddIn.ExcelApplication.Cursor = XlMousePointer.xlWait;

                bool rebuildReferenceFrameDropDown = true;

                if (WWTManager.IsValidMachine(Common.Globals.TargetMachine.MachineIP.ToString(), true))
                {
                    this.currentWorkbookMap.RefreshLayers();
                }
                else
                {
                    // Synchronize WWT layers.
                    SyncOnWWTNotRunning();
                    rebuildReferenceFrameDropDown = false;
                }

                // Rebind UI.
                this.BuildAndBindLayerDetailsViewModel(rebuildReferenceFrameDropDown, false);
            }
            finally
            {
                // Need to set the cursor to arrow, then default. Otherwise, busy cursor will shown until user moves to cursor away from where he clicked.
                ThisAddIn.ExcelApplication.Cursor = XlMousePointer.xlNorthwestArrow;
                ThisAddIn.ExcelApplication.Cursor = XlMousePointer.xlDefault;
            }
        }

        /// <summary>
        /// This function is called when the worksheet is deleted.
        /// </summary>
        private void OnSheetDelete()
        {
            // remove the affected layers from the workbook map
            WorkbookMap workbookMap = this.workBookMaps.Find(item => item.Workbook == this.mostRecentWorkbook);
            if (workbookMap != null)
            {
                workbookMap.RemoveAffectedLayers(this.affectedLayers);
            }

            this.mostRecentWorkbook.SaveWorkbookMap(this.workBookMaps);

            // Rebind UI
            this.BuildAndBindLayerDetailsViewModel();
        }

        /// <summary>
        /// Starts the call out animation
        /// </summary>
        private void BeginCalloutAnimation()
        {
            if (this.layerManagerPaneInstance != null)
            {
                Storyboard callOutstoryboard = (Storyboard)this.layerManagerPaneInstance.FindResource(Common.Constants.CallOutAnimation);
                callOutstoryboard.Begin();
            }
        }

        /// <summary>
        /// Updates the layer with the selected range
        /// </summary>
        private void UpdateLayer()
        {
            // Get Selection Range
            Range selectedRange = ThisAddIn.ExcelApplication.Selection as Range;

            if (this.currentWorkbookMap != null && this.currentWorkbookMap.SelectedLayerMap != null && selectedRange != null)
            {
                // Get Selected Layer.
                LayerMap selectedLayerMap = this.currentWorkbookMap.SelectedLayerMap;

                // Get named range name.
                string selectionRangeName = selectedLayerMap.RangeName.Name;

                // Remove NamedRange
                selectedLayerMap.RangeName.Delete();

                // Create Named Range
                Name updateNamedRange = this.currentWorkbookMap.Workbook.CreateNamedRange(selectionRangeName, selectedRange);

                // Reset named range properties.
                selectedLayerMap.ResetRange(updateNamedRange);

                // Update Header Data.
                selectedLayerMap.HeaderRowData = selectedRange.GetHeader();

                // Header Change
                // 1. AutoMap the columns 
                selectedLayerMap.SetAutoMap();

                // 2. Set layer properties dependent on mapping. 
                selectedLayerMap.SetLayerProperties();

                if (selectedLayerMap.MapType == LayerMapType.LocalInWWT && !selectedLayerMap.IsNotInSync)
                {
                    selectedLayerMap.IsNotInSync = true;
                }

                // Save the details of the workbook map to the custom xml parts
                ThisAddIn.ExcelApplication.ActiveWorkbook.SaveWorkbookMap(this.workBookMaps);

                // Update the View model to update in UI. 
                // No need to build the reference frame dropdown while updating the layer mapping.
                this.BuildAndBindLayerDetailsViewModel(false);
            }
        }

        /// <summary>
        /// Gets the layer map if it exists for the current range
        /// </summary>
        /// <param name="selectedRange">Selection range</param>
        /// <returns>Layer map if it exists for the selection range</returns>
        private LayerMap GetCurrentRangeLayer(Range selectedRange)
        {
            LayerMap layerMap = null;
            _Worksheet currentSheet = GetActiveWorksheet();
            if (selectedRange != null && this.currentWorkbookMap != null)
            {
                Range existingRange = null;

                // Get all Layer details into dictionary.
                foreach (LayerMap item in this.currentWorkbookMap.LocalLayerMaps)
                {
                    if (item.RangeName.IsWWTRange() && item.RangeName.IsValid() && item.RangeName.RefersToRange.IsValid()
                        && item.RangeName.RefersToRange.Worksheet.Equals(currentSheet))
                    {
                        existingRange = currentSheet.Application.Range[item.RangeName.Name];
                        if (existingRange.Areas.Count == selectedRange.Areas.Count)
                        {
                            int areaMatchCount = 0;
                            foreach (Range area in existingRange.Areas)
                            {
                                foreach (Range selectedArea in selectedRange.Areas)
                                {
                                    if (selectedArea.Address.Equals(area.Address))
                                    {
                                        areaMatchCount++;
                                        break;
                                    }
                                }
                            }
                            if (selectedRange.Areas.Count == areaMatchCount)
                            {
                                layerMap = item;
                                break;
                            }
                        }
                    }
                }
            }
            return layerMap;
        }
        
        #endregion Private methods

        #region Workbook Events

        /// <summary>
        /// When a protected view workbook is activated, disable the ribbon controls.
        /// </summary>
        /// <param name="protectedViewWindow">Protected view window</param>
        private void OnProtectedWorkbookActivate(ProtectedViewWindow protectedViewWindow)
        {
            // This is a protected workbook, just disable the ribbon and do nothing.
            this.ribbonInstance.EnableRibbonControls(false);
        }

        private void OnWorkbookActivate(Workbook workbook)
        {
            if (this.ribbonInstance != null)
            {
                // Enable the ribbon controls when we have one or more valid workbooks.
                // There could be scenario where excel is opening protected sheet.
                this.ribbonInstance.EnableRibbonControls(ThisAddIn.ExcelApplication.Workbooks.Count > 0);
            }

            var workbookMap = this.workBookMaps.Find(item => item.Workbook == workbook);
            if (workbookMap != null)
            {
                this.currentWorkbookMap = workbookMap;

                // No need to build the reference frame dropdown while workbook is getting activated.
                this.BuildAndBindLayerDetailsViewModel(false);
            }

            // Order in which EnableRibbonControls and BuildViewpointMenu are called is important
            var viewpointMap = this.viewpointMaps.Find(item => item.Workbook == workbook);
            if (viewpointMap != null && this.ribbonInstance != null)
            {
                this.currentViewpointMap = viewpointMap;
                this.ribbonInstance.BuildViewpointMenu(this.currentViewpointMap.SerializablePerspective);
            }

            // Update worksheet name.
            currentWorksheet = (Worksheet)ThisAddIn.ExcelApplication.ActiveSheet;
            worksheetName = currentWorksheet.Name;
        }

        private void OnNewWorkbook(Workbook workbook)
        {
            // Create a new workbookMap and add it to the list
            this.currentWorkbookMap = workbook.GetWorkbookMap();

            this.currentViewpointMap = workbook.GetViewpointMap();

            // Add this to the workbook map list
            this.workBookMaps.Add(this.currentWorkbookMap);

            this.viewpointMaps.Add(this.currentViewpointMap);
        }

        private void OnWorkbookOpen(Workbook workbook)
        {
            var workbookMap = this.workBookMaps.Find(item => item.Workbook == workbook);

            // This is for new condition else it is a reopen scenario 
            if (workbookMap == null)
            {
                this.currentWorkbookMap = workbook.GetWorkbookMap();

                // Add this to the workbook map list
                this.workBookMaps.Add(this.currentWorkbookMap);
            }

            var viewpointMap = this.viewpointMaps.Find(item => item.Workbook == workbook);

            // This is for new condition else it is a reopen scenario 
            if (viewpointMap == null)
            {
                this.currentViewpointMap = workbook.GetViewpointMap();

                // Add this to the viewpoint map list
                this.viewpointMaps.Add(this.currentViewpointMap);
            }

            if (this.ribbonInstance != null && this.currentWorkbookMap.LocalLayerMaps.Count > 0)
            {
                this.ribbonInstance.ViewCustomTaskPane(true);
            }

            // Sync up in case some workbook in excel was overwritten by open 
            SyncWorkbookMapWithExcel();
            SyncViewpointMapWithExcel();
        }

        private void OnWindowDeactivate(Workbook workbook, Window window)
        {
            // If the count is 1, then this workbook is last and is getting closed
            if (ThisAddIn.ExcelApplication.Workbooks.Count == 1)
            {
                if (this.ribbonInstance != null)
                {
                    // Disable all controls
                    this.ribbonInstance.EnableRibbonControls(false);
                }

                // Remove it from the list
                var workbookMap = this.workBookMaps.Find(item => item.Workbook == workbook);
                if (workbookMap != null)
                {
                    workbookMap.StopAllNotifications();
                    this.workBookMaps.Remove(workbookMap);
                }

                // Reset current workbook map as well
                this.currentWorkbookMap = null;

                // Remove it from the list
                var viewpointMap = this.viewpointMaps.Find(item => item.Workbook == workbook);
                if (viewpointMap != null)
                {
                    this.viewpointMaps.Remove(viewpointMap);
                }

                // Reset current workbook map as well
                this.currentWorkbookMap = null;
                this.currentViewpointMap = null;

                SyncWorkbookMapWithExcel();
                SyncViewpointMapWithExcel();

                if (this.ribbonInstance != null)
                {
                    this.ribbonInstance.ViewCustomTaskPane(false);
                }
            }
        }

        private void OnSheetChange(object sheet, Range target)
        {
            Worksheet currentSheet = sheet as Worksheet;

            if (currentSheet != null && target != null)
            {
                try
                {
                    // Update the details in CTP, WWT and Object models.
                    UpdateDetails(target, currentSheet);
                }
                catch (OutOfMemoryException ex)
                {
                    // consume this exception
                    Logger.LogException(ex);
                }
                finally
                {
                    if (afterCalculateCalled)
                    {
                        // This fix is specific to performance issue which happens with Find and Replace All.
                        // handledNamedRanges collection will make sure that even if SheetChange event is called multiple times
                        // for a single operation like Find and Replace All, UpdateData and BuildAndBindLayerDetailsViewModel
                        // will not be called unnecessarily.
                        // Only when a cell's value is changed by user manually, AfterCalculate will be called first. In that case, 
                        // dictionary object handledNamedRanges needs to be cleared, so that Layer to which the cell belongs to will be 
                        // updated properly.
                        // In all other cases, SheetChange will be called first and then final call will be AfterCalculate where
                        // handledNamedRanges will be cleared to make sure next SheetChange event will update the Layers properly.
                        handledNamedRanges.Clear();
                        afterCalculateCalled = false;
                    }
                }
            }
        }

        /// <summary>
        /// Check if the most recently used worksheet still exists
        /// </summary>
        /// <param name="sheet">worksheet object</param>
        private void OnSheetActivate(object sheet)
        {
            currentWorksheet = (Worksheet)ThisAddIn.ExcelApplication.ActiveSheet;
            worksheetName = currentWorksheet.Name;

            // Proceed only if activate has been fired immediately following a deactivate
            if (this.mostRecentWorkbook != null && this.mostRecentWorksheet != null)
            {
                bool sheetDeleted = true;
                foreach (Worksheet worksheet in this.mostRecentWorkbook.Worksheets)
                {
                    if (worksheet == this.mostRecentWorksheet)
                    {
                        // worksheet found 
                        sheetDeleted = false;
                        break;
                    }
                }

                // Call a method such as the following to execute code when you know that worksheet has been deleted
                if (sheetDeleted)
                {
                    this.OnSheetDelete();
                }

                this.mostRecentWorkbook = null;
                this.mostRecentWorksheet = null;
            }
        }

        /// <summary>
        /// Store the name of the sheet that is getting deactivated
        /// </summary>
        /// <param name="sheet">worksheet object</param>
        private void OnSheetDeactivate(object sheet)
        {
            if (sheet != null)
            {
                this.mostRecentWorksheet = sheet as Worksheet;
                this.mostRecentWorkbook = mostRecentWorksheet.Parent as Workbook;
                this.affectedLayers.Clear();
                var workbookMap = this.workBookMaps.Find(item => item.Workbook == this.mostRecentWorkbook);
                if (workbookMap != null)
                {
                    foreach (LayerMap layerMap in workbookMap.AllLayerMaps)
                    {
                        if (IsLocalLayer(layerMap.MapType) && (layerMap.RangeName != null && layerMap.RangeName.IsValid() && layerMap.RangeName.RefersToRange.Worksheet == this.mostRecentWorksheet))
                        {
                            this.affectedLayers.Add(layerMap);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// This function is called on almost every action in worksheet.
        /// </summary>
        private void OnAfterCalculate()
        {
            // Check if the current worksheet name has changed.
            if (worksheetName != currentWorksheet.Name && this.currentWorkbookMap != null)
            {
                foreach (LayerMap layer in this.currentWorkbookMap.LocalLayerMaps)
                {
                    // Check if the layer belongs to the current worksheet.
                    if (layer.RangeAddress.StartsWith(string.Format(System.Globalization.CultureInfo.InvariantCulture, "={0}!", worksheetName), StringComparison.Ordinal))
                    {
                        // if yes, then update the RangeAddress of all layer which belongs to the current worksheet.
                        layer.RangeAddress = layer.RangeAddress.Replace(worksheetName, currentWorksheet.Name);
                    }
                }
            }

            // Update the current worksheet details.
            currentWorksheet = (Worksheet)ThisAddIn.ExcelApplication.ActiveSheet;
            worksheetName = currentWorksheet.Name;

            afterCalculateCalled = true;
            handledNamedRanges.Clear();
        }

        #endregion Workbook Events

        #region Ribbon Events

        /// <summary>
        /// Event is raised when the Visualize selection button is clicked.
        /// </summary>
        /// <param name="sender">Visualize button</param>
        /// <param name="e">Routed even</param>
        private void OnVisualizeSelectionClicked(object sender, EventArgs e)
        {
            CreateLayerMap();
        }

        /// <summary>
        /// Event is raised when the target machine is changed.
        /// </summary>
        /// <param name="sender">Ribbon control</param>
        /// <param name="e">Routed event</param>
        private void OnTargetMachineChanged(object sender, EventArgs e)
        {
            if (this.currentWorkbookMap != null)
            {
                // Loads the WWT layers from the remote machine
                this.currentWorkbookMap.LoadWWTLayers();

                // Rebuilds the view model 
                this.BuildAndBindLayerDetailsViewModel();
            }
        }

        /// <summary>
        /// Event is raised when Capture Viewpoint is clicked.
        /// </summary>
        /// <param name="sender">Ribbon control</param>
        /// <param name="e">Routed event</param>
        private void OnGetViewpointClicked(object sender, EventArgs e)
        {
            Perspective perspective = sender as Perspective;
            if (perspective != null)
            {
                this.currentViewpointMap.SerializablePerspective.Add(perspective);
                SaveViewpointMap(ThisAddIn.ExcelApplication.ActiveWorkbook);
                this.ribbonInstance.BuildViewpointMenu(this.currentViewpointMap.SerializablePerspective);
            }
        }

        /// <summary>
        /// Event is raised when go to Viewpoint is clicked.
        /// </summary>
        /// <param name="sender">Ribbon control</param>
        /// <param name="e">Routed event</param>
        private void OnGotoViewpointClicked(object sender, EventArgs e)
        {
            Perspective perspective = sender as Perspective;
            if (perspective != null)
            {
                GotoViewpoint(perspective);
            }
        }

        /// <summary>
        /// Event is raised when go to Viewpoint from data is clicked.
        /// </summary>
        /// <param name="sender">Ribbon control</param>
        /// <param name="e">Routed event</param>
        private void OnGotoViewpointFromDataClicked(object sender, EventArgs e)
        {
            try
            {
                Utility.IsWWTInstalled();
                WWTManager.IsValidMachine(Common.Globals.TargetMachine.MachineIP.ToString(), false);
                var allNamedRange = this.currentWorkbookMap.GetNamedRangesForInSyncLayers();
                if (allNamedRange.Count > 0)
                {
                    _Worksheet workSheet = GetActiveWorksheet();

                    // Get range name from active cell
                    var rangeName = workSheet.GetRangeNameForActiveCell(ThisAddIn.ExcelApplication.ActiveCell, allNamedRange);
                    if (string.IsNullOrWhiteSpace(rangeName))
                    {
                        Ribbon.ShowError(Resources.GotoViewpointFromDataSelectionError);
                    }
                    else
                    {
                        LayerMap layerMap = this.currentWorkbookMap.LocalInWWTLayerMaps.Where(item => (!item.IsNotInSync) && item.RangeDisplayName.Equals(rangeName)).FirstOrDefault();
                        if (layerMap != null)
                        {
                            // Check if Lat/Lon, RA/Dec are mapped or not
                            if (!((layerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.Lat).Any()
                                    && layerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.Long).Any())
                                || (layerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.RA).Any()
                                    && layerMap.MappedColumnType.Where(columnTypeValue => columnTypeValue == ColumnType.Dec).Any())))
                            {
                                ShowViewpointFromDataValidationError(layerMap);
                            }
                            else
                            {
                                // Intersect active cell's row with the layer's range to get the layer row range
                                Range row = workSheet.Application.Intersect(workSheet.Application.Range[layerMap.RangeName.RefersTo as string] as Range, ThisAddIn.ExcelApplication.ActiveCell.EntireRow).Cells;

                                // Get the row data from the row range
                                var rowData = row.GetHeader();

                                // Build perspective based on mappings
                                var perspective = GetPerspectiveFromLayerRowData(layerMap, rowData);
                                if (perspective != null)
                                {
                                    GotoViewpointFromData(perspective);
                                }
                            }
                        }
                    }
                }
                else
                {
                    Ribbon.ShowError(Resources.GotoViewpointFromDataSelectionError);
                }
            }
            catch (CustomException ex)
            {
                Ribbon.ShowError(ex.HasCustomMessage ? ex.Message : Resources.LayerOperationError);
            }
        }

        /// <summary>
        /// Event is raised when manage Viewpoint is clicked.
        /// </summary>
        /// <param name="sender">Ribbon control</param>
        /// <param name="e">Routed event</param>
        private void OnManageViewpointClicked(object sender, EventArgs e)
        {
            manageViewpointInstance = new ManageViewpoint();
            System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(manageViewpointInstance);
            helper.Owner = (IntPtr)ThisAddIn.ExcelApplication.Hwnd;

            var viewModel = new ManageViewpointViewModel(this.currentViewpointMap.SerializablePerspective);
            viewModel.DeleteViewpointEvent += new EventHandler(OnDeleteViewpointEvent);
            viewModel.RenameViewpointEvent += new EventHandler(OnRenameViewpointEvent);
            viewModel.GotoViewpointEvent += new EventHandler(OnGotoViewpointEvent);

            if (viewModel.AllViewpoint.Any())
            {
                if (viewModel.AllViewpoint.Where(item => item.IsSelected == true).FirstOrDefault() == null)
                {
                    viewModel.AllViewpoint.First().IsSelected = true;
                }

                viewModel.IsSelected = true;
            }
            else
            {
                viewModel.IsSelected = false;
            }

            manageViewpointInstance.DataContext = viewModel;
            manageViewpointInstance.ShowDialog();
        }

        /// <summary>
        /// Event is raised when Go to Viewpoint is clicked from the grid.
        /// </summary>
        /// <param name="sender">Ribbon control</param>
        /// <param name="e">Routed event</param>
        private void OnGotoViewpointEvent(object sender, EventArgs e)
        {
            ViewpointViewModel viewModel = sender as ViewpointViewModel;
            if (viewModel != null)
            {
                GotoViewpoint(viewModel.CurrentPerspective);
            }
        }

        /// <summary>
        /// Event is raised when rename Viewpoint is clicked from the grid.
        /// </summary>
        /// <param name="sender">Ribbon control</param>
        /// <param name="e">Routed event</param>
        private void OnRenameViewpointEvent(object sender, EventArgs e)
        {
            var viewModel = sender as ViewpointViewModel;
            if (viewModel != null)
            {
                CaptureViewpoint dialog = new CaptureViewpoint();
                System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(dialog);
                helper.Owner = (IntPtr)ThisAddIn.ExcelApplication.Hwnd;

                viewModel.RequestClose += new EventHandler(dialog.OnRequestClose);
                dialog.DataContext = viewModel;
                dialog.UpdateLayout();

                dialog.ShowDialog();

                if (dialog.DialogResult.HasValue && dialog.DialogResult.Value)
                {
                    if (!string.IsNullOrWhiteSpace(viewModel.Name))
                    {
                        viewModel.CurrentPerspective.Name = viewModel.Name;
                        SaveViewpointMap(ThisAddIn.ExcelApplication.ActiveWorkbook);
                        this.ribbonInstance.BuildViewpointMenu(this.currentViewpointMap.SerializablePerspective);
                    }
                }

                viewModel.RequestClose -= new EventHandler(dialog.OnRequestClose);
                dialog.Close();
            }
        }

        /// <summary>
        /// Event is raised when delete Viewpoint is clicked from the grid.
        /// </summary>
        /// <param name="sender">Ribbon control</param>
        /// <param name="e">Routed event</param>
        private void OnDeleteViewpointEvent(object sender, EventArgs e)
        {
            var viewpointViewModel = sender as ViewpointViewModel;
            if (viewpointViewModel != null)
            {
                var perspective = viewpointViewModel.CurrentPerspective;

                // Remove and save the entries
                this.currentViewpointMap.SerializablePerspective.Remove(perspective);
                SaveViewpointMap(ThisAddIn.ExcelApplication.ActiveWorkbook);

                // Rebuild the model from the remaining set and rebind it
                var viewModel = new ManageViewpointViewModel(this.currentViewpointMap.SerializablePerspective);
                viewModel.DeleteViewpointEvent += new EventHandler(OnDeleteViewpointEvent);
                viewModel.RenameViewpointEvent += new EventHandler(OnRenameViewpointEvent);
                viewModel.GotoViewpointEvent += new EventHandler(OnGotoViewpointEvent);

                manageViewpointInstance.DataContext = viewModel;
                if (viewModel.AllViewpoint.Any())
                {
                    if (viewModel.AllViewpoint.Where(item => item.IsSelected == true).FirstOrDefault() == null)
                    {
                        viewModel.AllViewpoint.First().IsSelected = true;
                    }

                    viewModel.IsSelected = true;
                }
                else
                {
                    viewModel.IsSelected = false;
                }

                this.manageViewpointInstance.UpdateLayout();
                this.ribbonInstance.BuildViewpointMenu(this.currentViewpointMap.SerializablePerspective);
            }
        }

        #endregion Ribbon Events

        #region CustomTaskPane Events

        /// <summary>
        /// When in Excel is SDI, the pane-workbook collection is managed by 'LayerTaskPaneController'.
        /// When active workbook changes, corresponding task pane changes. Workflow controller needs to be refering to current task pane
        /// </summary>
        void OnLayerPaneChangedEvent(LayerManagerPane pane)
        {
            UpdateLayerManagerPaneInstance(pane);
        }

        /// <summary>
        /// Event is raised on the state change of custom task pane
        /// </summary>
        /// <param name="sender">Layer details view model</param>
        /// <param name="e">Routed event</param>
        private void OnCustomTaskPaneChangedState(object sender, EventArgs e)
        {
            if (this.currentWorkbookMap != null && this.currentWorkbookMap.SelectedLayerMap != null
                && this.currentWorkbookMap.SelectedLayerMap.LayerDetails != null && this.layerDetailsViewModel != null)
            {
                // Set the layer map name.
                this.currentWorkbookMap.SelectedLayerMap.LayerDetails.Name = this.layerDetailsViewModel.SelectedLayerName;

                // Set map column values to object model
                this.currentWorkbookMap.SelectedLayerMap.MappedColumnType.Clear();
                foreach (ColumnViewModel columnView in this.layerDetailsViewModel.ColumnsView)
                {
                    this.currentWorkbookMap.SelectedLayerMap.MappedColumnType.Add(columnView.SelectedWWTColumn.ColType);
                }

                // Set the group.
                this.currentWorkbookMap.SelectedLayerMap.LayerDetails.Group = this.layerDetailsViewModel.SelectedGroup;

                // Set the indices of Layer Details column properties based on latest mapping
                this.currentWorkbookMap.SelectedLayerMap.SetLayerColumnProperties();

                // Set the last used group to the selected group.
                LastUsedGroup = this.layerDetailsViewModel.SelectedGroup;

                // Sets the selected distance column if selected
                if (this.layerDetailsViewModel.IsDistanceVisible)
                {
                    this.currentWorkbookMap.SelectedLayerMap.LayerDetails.AltUnit = this.layerDetailsViewModel.SelectedDistanceUnit.Key;
                }

                // Sets the selected RA column if selected
                if (this.layerDetailsViewModel.IsRAUnitVisible)
                {
                    this.currentWorkbookMap.SelectedLayerMap.LayerDetails.RAUnit = this.layerDetailsViewModel.SelectedRAUnit.Key;
                }

                this.currentWorkbookMap.SelectedLayerMap.LayerDetails.FadeType = this.layerDetailsViewModel.SelectedFadeType.Key;
                this.currentWorkbookMap.SelectedLayerMap.LayerDetails.PointScaleType = this.layerDetailsViewModel.SelectedScaleType.Key;
                this.currentWorkbookMap.SelectedLayerMap.LayerDetails.MarkerScale = this.layerDetailsViewModel.SelectedScaleRelative.Key;

                this.currentWorkbookMap.SelectedLayerMap.LayerDetails.Opacity = this.layerDetailsViewModel.LayerOpacity.SelectedSliderValue / 100;
                this.currentWorkbookMap.SelectedLayerMap.LayerDetails.TimeDecay = this.layerDetailsViewModel.GetActualTimeDecayValue();
                this.currentWorkbookMap.SelectedLayerMap.LayerDetails.ScaleFactor = this.layerDetailsViewModel.GetActualScaleFactorValue();

                this.currentWorkbookMap.SelectedLayerMap.LayerDetails.SizeColumn = this.layerDetailsViewModel.SelectedSize.Key;
                this.currentWorkbookMap.SelectedLayerMap.LayerDetails.NameColumn = this.layerDetailsViewModel.SelectedHoverText.Key;

                this.currentWorkbookMap.SelectedLayerMap.LayerDetails.ShowFarSide = this.layerDetailsViewModel.IsFarSideShown;
                this.currentWorkbookMap.SelectedLayerMap.LayerDetails.PlotType = this.layerDetailsViewModel.SelectedMarkerType.Key;
                this.currentWorkbookMap.SelectedLayerMap.LayerDetails.MarkerIndex = this.layerDetailsViewModel.SelectedPushpinId.Key;

                // Set the Mapped column type based on the group selected.
                this.currentWorkbookMap.SelectedLayerMap.UpdateMappedColumns();

                SetCoordinateType();

                // Any change in Core OM needs to be saved
                ThisAddIn.ExcelApplication.ActiveWorkbook.SaveWorkbookMap(this.workBookMaps);

                this.UpdateWWTPropertiesForSelectedLayer();
            }
        }

        /// <summary>
        /// Event is raised on the state change of layer drop down
        /// </summary>
        /// <param name="sender">LayerMapDropDown ViewModel</param>
        /// <param name="e">Routed event</param>
        private void OnLayerSelectionChanged(object sender, EventArgs e)
        {
            LayerMap selectedLayerMap = sender as LayerMap;

            // The callout is required in the scenario when the layer is selected from the dropdown.
            LayerDetailsViewModel.IsCallOutRequired = true;
            if (this.currentWorkbookMap != null && selectedLayerMap != null)
            {
                this.currentWorkbookMap.SelectedLayerMap = GetSelectedLayerMap(selectedLayerMap);
                this.layerDetailsViewModel.Currentlayer = this.currentWorkbookMap.SelectedLayerMap;
                this.layerDetailsViewModel.IsReferenceGroupEnabled = this.currentWorkbookMap.SelectedLayerMap.IsLayerCreated();
                SetGetLayerDataDisplayName(this.currentWorkbookMap.SelectedLayerMap);

                // Any change in Core OM needs to be saved
                ThisAddIn.ExcelApplication.ActiveWorkbook.SaveWorkbookMap(this.workBookMaps);

                if (IsLocalLayer(selectedLayerMap.MapType))
                {
                    // Show the range selected.
                    this.ShowSelectedRange();
                }
            }
            else
            {
                // If the layers are blank the selected layer set to null 
                this.currentWorkbookMap.SelectedLayerMap = null;

                // Any change in Core OM needs to be saved
                ThisAddIn.ExcelApplication.ActiveWorkbook.SaveWorkbookMap(this.workBookMaps);

                // Rebind UI
                this.BuildAndBindLayerDetailsViewModel();
            }
        }

        /// <summary>
        /// Event is raised on the click of View in WWT button
        /// </summary>
        /// <param name="sender">View in WWT button.</param>
        /// <param name="e">Routed event</param>
        private void OnViewInWWTClicked(object sender, EventArgs e)
        {
            ViewInWWT();
        }

        /// <summary>
        /// Event is raised on the show range click event.
        /// </summary>
        /// <param name="sender">
        /// Show Range button.
        /// </param>
        /// <param name="e">
        /// Routed event
        /// </param>
        private void OnShowRangeClickedEvent(object sender, EventArgs e)
        {
            this.ShowSelectedRange();
        }

        /// <summary>
        /// Event is raised on Layer drop down Open event.
        /// </summary>
        /// <param name="sender">
        /// Layer drop down.
        /// </param>
        /// <param name="e">
        /// Routed event
        /// </param>
        private void OnRefreshDropDownClickedEvent(object sender, EventArgs e)
        {
            this.OnRefreshMapping();
        }

        /// <summary>
        /// Event is raised on the delete mapping click operation.
        /// </summary>
        /// <param name="sender">
        /// Delete mapping button.
        /// </param>
        /// <param name="e">
        /// Routed event
        /// </param>
        private void OnDeleteMappingClickedEvent(object sender, EventArgs e)
        {
            this.DeleteMapping();
        }

        /// <summary>
        /// Event is raised on get layer data click operation
        /// </summary>
        /// <param name="sender">Get layer data button</param>
        /// <param name="e">Routed event</param>
        private void OnGetLayerDataClickedEvent(object sender, EventArgs e)
        {
            try
            {
                WWTManager.IsValidMachine(Common.Globals.TargetMachine.MachineIP.ToString(), false);
                this.GetLayerData();
            }
            catch (CustomException ex)
            {
                this.SyncAndRebindOnError(ex);
            }
        }

        /// <summary>
        /// Event is raised on the Update layer click event.
        /// This function is used to validate the selection range and update layer if valid
        /// </summary>
        /// <param name="sender">
        /// Update layer button.
        /// </param>
        /// <param name="e">
        /// Routed event
        /// </param>
        private void OnUpdateLayerClickedEvent(object sender, EventArgs e)
        {
            Range selectedRange = ThisAddIn.ExcelApplication.Selection as Range;
            if (selectedRange != null && selectedRange.IsValid())
            {
                this.UpdateLayer();
            }
            else
            {
                Ribbon.ShowError(Properties.Resources.InvalidSelectionRange);
            }
        }

        /// <summary>
        /// Refreshes the reference frame dropdown and the layer details view model
        /// </summary>
        /// <param name="sender">Custom task pane</param>
        /// <param name="e">Routed event</param>
        private void OnRefreshGroupDropDownClickedEvent(object sender, EventArgs e)
        {
            bool isValidMachine = true;
            CustomException exception = null;

            try
            {
                ThisAddIn.ExcelApplication.Cursor = XlMousePointer.xlWait;

                try
                {
                    isValidMachine = WWTManager.IsValidMachine(Common.Globals.TargetMachine.MachineIP.ToString(), false);
                }
                catch (CustomException ex)
                {
                    // Both exceptions, WWT not open and WWT Version mismatch will be handled here.
                    isValidMachine = false;
                    exception = ex;
                }

                if (!isValidMachine)
                {
                    // Synchronize WWT layers.
                    SyncOnWWTNotRunning();
                }

                // Rebind UI.
                this.BuildAndBindLayerDetailsViewModel(isValidMachine);
            }
            finally
            {
                // Need to set the cursor to arrow, then default. Otherwise, busy cursor will shown until user moves to cursor away from where he clicked.
                ThisAddIn.ExcelApplication.Cursor = XlMousePointer.xlNorthwestArrow;
                ThisAddIn.ExcelApplication.Cursor = XlMousePointer.xlDefault;
            }

            if (exception != null)
            {
                // Error message needs to be shown here, otherwise busy cursor will be shown while error message is shown.
                Ribbon.ShowError(exception.Message);
            }
        }

        /// <summary>
        /// Reference frame selection change from sky to planet or vice-versa
        /// will auto-map the columns
        /// </summary>
        /// <param name="sender">Custom task pane</param>
        /// <param name="e">Routed event</param>
        private void OnReferenceSelectionChanged(object sender, EventArgs e)
        {
            if (this.currentWorkbookMap != null && this.currentWorkbookMap.SelectedLayerMap != null)
            {
                // Sets auto map for columns and update the object model
                layerDetailsViewModel.Currentlayer.SetAutoMap();
                this.currentWorkbookMap.SelectedLayerMap = layerDetailsViewModel.Currentlayer;
            }
        }

        /// <summary>
        /// Handle the click event for the download updates button on the custom task pane
        /// </summary>
        /// <param name="sender">Download Updates button</param>
        /// <param name="e">Default Event Arguments</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        private void OnDownloadUpdatesClicked(object sender, EventArgs e)
        {
            try
            {
                // start the download link worker
                this.updateManager.DownloadUpdates();

                // set the visibility and text of the download buttons
                this.downloadUpdatesViewModel.IsDownloadUpdatesEnabled = false;
                this.downloadUpdatesViewModel.DownloadUpdatesLabel = Resources.DownloadUpdatesButtonDownloadingLabel;

                this.ribbonInstance.downloadUpdatesButton.Enabled = false;
                this.ribbonInstance.downloadUpdatesButton.Label = Resources.DownloadUpdatesButtonDownloadingLabel;
            }
            catch (CustomException ex)
            {
                Logger.LogException(ex);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
            }
        }
        #endregion

        #region UpdateManager Events

        /// <summary>
        /// Update Available event handler for the update manager
        /// </summary>
        /// <param name="sender">UpdateManager instance</param>
        /// <param name="e">Default Event Arguments</param>
        private void OnUpdateAvailable(object sender, EventArgs e)
        {
            // set the visibility of the buttons on the task pane and ribbon
            this.downloadUpdatesViewModel.IsDownloadUpdatesEnabled = true;
            this.downloadUpdatesViewModel.IsDownloadUpdatesVisible = true;
            this.ribbonInstance.updateGroup.Visible = true;
            this.ribbonInstance.downloadUpdatesButton.Enabled = true;
        }

        /// <summary>
        /// Download Completed event handler for the update manager
        /// </summary>
        /// <param name="sender">UpdateManager instance</param>
        /// <param name="e">Default Event Arguments</param>
        private void OnDownloadCompleted(object sender, EventArgs e)
        {
            try
            {
                // Reset the labels on the buttons 
                this.downloadUpdatesViewModel.DownloadUpdatesLabel = Resources.DownloadUpdatesButtonDownloadCompletedLabel;
                this.ribbonInstance.downloadUpdatesButton.Label = Resources.DownloadUpdatesButtonDownloadCompletedLabel;

                // Start the installer
                this.updateManager.InstallUpdates();
            }
            catch (System.ComponentModel.InvalidEnumArgumentException ex)
            {
                Logger.LogException(ex);
            }
            catch (System.InvalidOperationException ex)
            {
                Logger.LogException(ex);
            }
            catch (CustomException ex)
            {
                Logger.LogException(ex);
            }
        }

        /// <summary>
        /// Installation Completed event handler for the update manager
        /// </summary>
        /// <param name="sender">UpdateManager instance</param>
        /// <param name="e">Default Event Arguments</param>
        private void OnInstallationCompleted(object sender, EventArgs e)
        {
            // Reset the labels on the buttons and enable them
            this.downloadUpdatesViewModel.DownloadUpdatesLabel = Resources.DownloadUpdatesButtonLabel;
            this.ribbonInstance.downloadUpdatesButton.Label = Resources.DownloadUpdatesButtonLabel;

            this.downloadUpdatesViewModel.IsDownloadUpdatesEnabled = true;
            this.ribbonInstance.downloadUpdatesButton.Enabled = true;
        }

        #endregion UpdateManager Events
    }
}
