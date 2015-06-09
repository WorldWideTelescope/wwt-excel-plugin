//-----------------------------------------------------------------------
// <copyright file="Ribbon.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Research.Wwt.Excel.Addin.Properties;
using Microsoft.Research.Wwt.Excel.Common;
using Microsoft.Win32;
using System.Collections.Generic;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// This class is the UI for the WWT ribbon in Excel.
    /// </summary>
    public partial class Ribbon
    {
        /// <summary>
        /// Name of the group on the ribbon that contains the download updates button
        /// </summary>
        private const string RibbonDownloadUpdatesGroupName = "updateGroup";

        /// <summary>
        /// Context menu button.
        /// </summary>
        private CommandBarButton cellVisualizeMenu;

        /// <summary>
        /// VisualizeSelectionClicked event
        /// </summary>
        internal event EventHandler VisualizeSelectionClicked;

        /// <summary>
        /// GetViewpointClicked event
        /// </summary>
        internal event EventHandler GetViewpointClicked;

        /// <summary>
        /// GotoViewpointClicked event
        /// </summary>
        internal event EventHandler GotoViewpointClicked;

        /// <summary>
        /// GotoViewpointFromDataClicked event
        /// </summary>
        internal event EventHandler GotoViewpointFromDataClicked;

        /// <summary>
        /// ManageViewpointClicked event
        /// </summary>
        internal event EventHandler ManageViewpointClicked;

        /// <summary>
        /// TargetMachineChanged event.
        /// </summary>
        internal event EventHandler TargetMachineChanged;

        /// <summary>
        /// Download updates button clicked event
        /// </summary>
        internal event EventHandler DownloadUpdatesButtonClicked;

        /// <summary>
        /// Gets the value indicating whether the auto move is enabled or not.
        /// </summary>
        internal bool IsAutoMoveEnabled
        {
            get
            {
                return this.tgBtnAutoMove.Checked;
            }
        }

        #region Internal & public methods

        /// <summary>
        /// Displays a message box for Error scenario
        /// </summary>
        /// <param name="message">The message to be displayed</param>
        internal static void ShowError(string message)
        {
            // This is needed so that unit test cases will not be blocked in scenarios where error messages are shown.
            if (Globals.ThisAddIn != null)
            {
                MessageBox.Show(ThisAddIn.ExcelApplication.ActiveWindow as IWin32Window, message, Resources.MessageBoxCaption, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }
        }

        /// <summary>
        /// Displays a message box for warning scenario
        /// </summary>
        /// <param name="message">The message to be displayed</param>
        internal static void ShowWarning(string message)
        {
            // This is needed so that unit test cases will not be blocked in scenarios where warning messages are shown.
            if (Globals.ThisAddIn != null)
            {
                MessageBox.Show(ThisAddIn.ExcelApplication.ActiveWindow as IWin32Window, message, Resources.MessageBoxCaption, MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
            }
        }

        /// <summary>
        /// Displays a message box for warning scenario with YEs No Button and return the user selected option.
        /// </summary>
        /// <param name="message">The message to be displayed</param>
        /// <returns>True if the user has selected Yes;Otherwise false.</returns>
        internal static bool ShowWarningWithResult(string message)
        {
            DialogResult result = MessageBox.Show(
                                                ThisAddIn.ExcelApplication.ActiveWindow as IWin32Window,
                                                message,
                                                Properties.Resources.MessageBoxCaption,
                                                MessageBoxButtons.YesNo,
                                                MessageBoxIcon.Warning,
                                                MessageBoxDefaultButton.Button1);

            return result == DialogResult.Yes;
        }

        /// <summary>
        /// Enable or disable the ribbon controls on the Tab.
        /// </summary>
        /// <param name="value">True to enable the all the controls, otherwise false.</param>
        internal void EnableRibbonControls(bool value)
        {
            RibbonTab tab = this.wwteRibbonTab;
            foreach (RibbonGroup group in tab.Groups)
            {
                if (group.Id == Ribbon.RibbonDownloadUpdatesGroupName)
                {
                    // If the group is Update, its state will be enabled/disabled by the event handlers.
                    continue;
                }

                foreach (RibbonControl control in group.Items)
                {
                    control.Enabled = value;
                }
            }

            // The gallery needs to be enabled all the time
            this.viewSamplesGallery.Enabled = true;
        }

        /// <summary>
        /// Hide or show custom task pane
        /// </summary>
        /// <param name="isVisible">Is visible</param>
        internal void ViewCustomTaskPane(bool isVisible)
        {
            if (LayerTaskPaneController.Instance.CurrentTaskPane != null)
                LayerTaskPaneController.Instance.CurrentTaskPane.Visible = isVisible;
        }

        /// <summary>
        /// Sets the focus on custom task pane
        /// </summary>
        internal void SetFocusCustomTaskPane()
        {
            if (LayerTaskPaneController.Instance.CurrentTaskPane != null)
                LayerTaskPaneController.Instance.CurrentTaskPane.Control.Focus();
        }

        /// <summary>
        /// Builds view point menu list
        /// </summary>
        /// <param name="viewpointList">perspective List items</param>
        internal void BuildViewpointMenu(ObservableCollection<Perspective> viewpointList)
        {
            this.gotoViewpointGallery.Items.Clear();
            foreach (Perspective viewpoint in viewpointList)
            {
                var viewpointMenuItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                viewpointMenuItem.Label = viewpoint.Name;
                viewpointMenuItem.Tag = viewpoint;
                this.gotoViewpointGallery.Items.Add(viewpointMenuItem);
            }

            if (this.gotoViewpointGallery.Items.Count == 0)
            {
                this.gotoViewpointGallery.Enabled = false;
            }
            else
            {
                this.gotoViewpointGallery.Enabled = true;
            }
        }

        #endregion

        #region Private static methods

        /// <summary>
        /// This method reads the registry to obtain the application installation location.
        /// </summary>
        /// <returns>String - app installation location.</returns>
        private static string GetAppInstallationLocation()
        {
            string installationLocation = string.Empty;
            try
            {
                // Get the Registry key specific to the Local Machine. 
                RegistryKey regKey;
                using (var regKeyHkCU = Registry.LocalMachine)
                {
                    regKey = regKeyHkCU.OpenSubKey(Common.Constants.AppRegistryPath);
                }

                // Access the key which gives the installation location.
                if (regKey != null && regKey.GetValue(Common.Constants.AppRegistryKey) != null)
                {
                    var manifestLocation = regKey.GetValue(Common.Constants.AppRegistryKey).ToString();
                    if (!string.IsNullOrEmpty(manifestLocation))
                    {
                        if (manifestLocation.Contains("|"))
                        {
                            manifestLocation = manifestLocation.Substring(0, manifestLocation.IndexOf("|", StringComparison.OrdinalIgnoreCase));
                        }

                        installationLocation = Path.GetDirectoryName(manifestLocation);
                    }
                }
            }
            catch (System.Security.SecurityException ex)
            {
                Logger.LogException(ex);
                throw new CustomException(String.Format(CultureInfo.InvariantCulture, Resources.ErrorReadingRegistry, Properties.Resources.ProductNameShort), true);
            }
            catch (ArgumentException ex)
            {
                Logger.LogException(ex);
                throw new CustomException(String.Format(CultureInfo.InvariantCulture, Resources.ErrorReadingRegistry, Properties.Resources.ProductNameShort), true);
            }
            catch (UnauthorizedAccessException ex)
            {
                Logger.LogException(ex);
                throw new CustomException(String.Format(CultureInfo.InvariantCulture, Resources.ErrorReadingRegistry, Properties.Resources.ProductNameShort), true);
            }
            catch (IOException ex)
            {
                Logger.LogException(ex);
                throw new CustomException(String.Format(CultureInfo.InvariantCulture, Resources.ErrorReadingRegistry, Properties.Resources.ProductNameShort), true);
            }

            return installationLocation;
        }

        /// <summary>
        /// Start process based on target
        /// </summary>
        /// <param name="target">target name</param>
        private static void StartProcess(string target)
        {
            try
            {
                Process.Start(target);
            }
            catch (Win32Exception)
            {
                Ribbon.ShowError(Resources.TryAgainErrorMessage);
            }
            catch (ObjectDisposedException)
            {
                Ribbon.ShowError(Resources.DefaultErrorMessage);
            }
            catch (FileNotFoundException)
            {
                Ribbon.ShowError(Resources.DefaultErrorMessage);
            }
        }

        #endregion

        #region Event handlers

        /// <summary>
        /// Occurs when the Ribbon is loaded into the Microsoft Office application. 
        /// </summary>
        /// <param name="sender">
        /// Event sender
        /// </param>
        /// <param name="e">
        /// Provides data for events that are raised in the Ribbon.
        /// </param>
        private void OnRibbonLoad(object sender, RibbonUIEventArgs e)
        {
            // Set the Excel application to Common.Globals, so that it can be used across.
            ThisAddIn.ExcelApplication = Globals.ThisAddIn.Application;

            //Initialize the Layer Manager TaskPaneController and let it handle the custom LayerManager taskpanes.
            LayerTaskPaneController.Instance.Initialize();

            //Link the LayerTaskPaneController which holds the references of multiple layer panes, with the workflowcontroller
            WorkflowController.Instance.RegisterForLayerPaneChangedEvent();

            // Set the WWT application path and exe name
            Utility.SetWWTApplicationPath();

            // Set target machine to local.
            Common.Globals.TargetMachine = new TargetMachine();

            if (ThisAddIn.ExcelApplication.Workbooks.Count > 0)
            {
                // Enable  the ribbon controls
                EnableRibbonControls(true);
            }
            else
            {
                // Disable the ribbon controls since the ribbon is loaded for Excel workbook which is not having any workbook.
                // This scenario will occur when excel is opening a protected workbook.
                EnableRibbonControls(false);
            }

            // Initialize the workflow controller
            if (LayerTaskPaneController.IsExcelInstanceSDI)
            {
                //Initialize with layermanagerpane to null. Workflowcontroller is initialized appropriately When a new/existing workbook is created/opened. 
                //see: LayerTaskPaneController - new workbook, open workbook & workbook activate event
                WorkflowController.Instance.Initialize(null, this);
            }
            else
            {
                WorkflowController.Instance.Initialize(LayerTaskPaneController.Instance.CurrentPaneHost.LayerManagerPane, this);
            }

            // Initializes the context menu for adding new button for "Visualize in WWT"
            InitializeContextControls();

            // Fill the defaults for the ribbon controls
            this.PopulateSamples();
            this.PopulateFeedbackGallery();

            // Looks for any network address change in the local machine.
            NetworkChange.NetworkAddressChanged += new NetworkAddressChangedEventHandler(OnNetworkAddressChanged);
        }

        /// <summary>
        /// Network address changed event handler which will be resetting the TargetMachine instance
        /// for the changed IP address so that the call to WWT LCAPI will go through smoothly.
        /// Resetting TargetMachine instance will happen only for local machine not for remote machine.
        /// </summary>
        /// <param name="sender">Null, since there is no sender</param>
        /// <param name="e">Empty event arguments</param>
        private void OnNetworkAddressChanged(object sender, EventArgs e)
        {
            // Only for local machine, need to reset the TargetMachine instance. Remote machine cannot be monitored.
            if (Common.Globals.TargetMachine.IsLocalMachine)
            {
                // Reset target machine with the changed IP address.
                Common.Globals.TargetMachine = new TargetMachine();
            }
        }

        /// <summary>
        /// This event handler fires on Close of the Ribbon.
        /// </summary>
        /// <param name="sender">sender object</param>
        /// <param name="e">event arguments</param>
        private void OnRibbonClose(object sender, EventArgs e)
        {
           
        }

        /// <summary>
        /// Add the sample files to the gallery on the ribbon
        /// </summary>
        private void PopulateSamples()
        {
            var sampleFileOneMenuItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            sampleFileOneMenuItem.Label = Resources.EarthBasedPointDataSampleLabel;
            sampleFileOneMenuItem.Tag = Resources.EarthBasedPointDataSampleFileName;
            sampleFileOneMenuItem.Image = Resources.Earth;
            var sampleFileTwoMenuItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            sampleFileTwoMenuItem.Label = Resources.EarthBasedGeometryDataSampleLabel;
            sampleFileTwoMenuItem.Tag = Resources.EarthBasedGeometryDataSampleFileName;
            sampleFileTwoMenuItem.Image = Resources.Earth;

            var sampleFileAstronomyOneMenuItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            sampleFileAstronomyOneMenuItem.Label = Resources.AstronomyBasedPointDataSampleLabel;
            sampleFileAstronomyOneMenuItem.Tag = Resources.SampleAstronomyPointBasedFile;
            sampleFileAstronomyOneMenuItem.Image = Resources.Sky;
            var sampleFileAstronomyTwoMenuItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            sampleFileAstronomyTwoMenuItem.Label = Resources.AstronomyBasedGeometryDataSampleLabel;
            sampleFileAstronomyTwoMenuItem.Tag = Resources.SampleAstronomyGeometryFile;
            sampleFileAstronomyTwoMenuItem.Image = Resources.Sky;

            this.viewSamplesGallery.Items.Add(sampleFileOneMenuItem);
            this.viewSamplesGallery.Items.Add(sampleFileTwoMenuItem);
            this.viewSamplesGallery.Items.Add(sampleFileAstronomyOneMenuItem);
            this.viewSamplesGallery.Items.Add(sampleFileAstronomyTwoMenuItem);
        }

        /// <summary>
        /// Add the links to the feedback gallery
        /// </summary>
        private void PopulateFeedbackGallery()
        {
            var surveyFeedbackItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            surveyFeedbackItem.Label = Resources.FeedbackSurveyLabel;
            surveyFeedbackItem.Image = Resources.Survey;
            var viewForumItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            viewForumItem.Label = Resources.VisitForumLabel;
            viewForumItem.Image = Resources.Forum;

            this.feedbackGallery.Items.Add(surveyFeedbackItem);
            this.feedbackGallery.Items.Add(viewForumItem);
        }

        /// <summary>
        /// Click event handler for the gallery items displaying sample files
        /// </summary>
        /// <param name="sender">Gallery on the ribbon</param>
        /// <param name="e">RibbonControlEventArgs instance</param>
        private void OnViewSamplesGalleryClick(object sender, RibbonControlEventArgs e)
        {
            

            RibbonGallery gallery = sender as RibbonGallery;
            if (gallery != null)
            {
                // find out which item was clicked
                // read the file name
                string fileName = (string)gallery.SelectedItem.Tag;

               
                try
                {
                    // find the file path
                    string sampleFilePath = Path.Combine(Ribbon.GetAppInstallationLocation(), fileName);

                    // open the file in the excel instance
                    if (File.Exists(sampleFilePath))
                    {
                        int wbCount = ThisAddIn.ExcelApplication.Workbooks.Count;
                        Workbook targetWb = ThisAddIn.ExcelApplication.ActiveWorkbook;
                        Workbook wb = ThisAddIn.ExcelApplication.Workbooks.Open(sampleFilePath, Type.Missing, true);

                        if(ThisAddIn.ExcelApplication.Workbooks.Count != wbCount)
                        {
                            Worksheet sourceWs = (Worksheet)wb.Worksheets[1];
                            Worksheet targetWs = (Worksheet)targetWb.Worksheets[1];
                            sourceWs.Copy(targetWs);
                            wb.Close(SaveChanges:false);
                        }
                    }
                    else
                    {
                        Ribbon.ShowError(Resources.SampleFileNotFound);
                    }
                }
                catch (CustomException ex)
                {
                    Ribbon.ShowError(ex.HasCustomMessage ? ex.Message : Resources.DefaultErrorMessage);
                }
                catch (COMException ex)
                {
                    Logger.LogException(ex);

                    // This is thrown when user says no in Re-open scenario
                    // Consume this exception
                }
            }
        }

        /// <summary>
        /// Click event handler for the feedback gallery
        /// </summary>
        /// <param name="sender">Feedback gallery</param>
        /// <param name="e">Ribbon control event arguments</param>
        private void OnFeedbackGalleryClick(object sender, RibbonControlEventArgs e)
        {
         
            RibbonGallery gallery = sender as RibbonGallery;
            if (gallery != null)
            {
                RibbonDropDownItem selectedItem = (RibbonDropDownItem)gallery.SelectedItem;
                if (selectedItem.Label == Resources.FeedbackSurveyLabel)
                {
                    StartProcess(Common.Constants.FeedbackLink);
                }
                else if (selectedItem.Label == Resources.VisitForumLabel)
                {
                    StartProcess(Common.Constants.VisitForumLink);
                }
            }
        }

        /// <summary>
        /// Handles layer manager 
        /// </summary>
        /// <param name="sender">Layer pane.</param>
        /// <param name="e">Ribbon Control Event arguments.</param>
        private void OnLayerButtonClick(object sender, RibbonControlEventArgs e)
        {
            if (LayerTaskPaneController.Instance.CurrentTaskPane != null)
                LayerTaskPaneController.Instance.CurrentTaskPane.Visible = LayerTaskPaneController.Instance.CurrentTaskPane.Visible ? false : true;
        }

        /// <summary>
        /// Handles visualize button click.
        /// </summary>
        /// <param name="sender">Visualize in WWT button.</param>
        /// <param name="e">Ribbon Control Event arguments.</param>
        private void OnVisualizeSelectionButtonClick(object sender, RibbonControlEventArgs e)
        {
            // Validate and create named Range.
            CreateNamedRange();
        }

        /// <summary>
        /// Occurs when the visualize in WWT button in cell context menu is clicked.
        /// </summary>
        /// <param name="ctrl">
        /// Event sender.
        /// </param>
        /// <param name="cancelDefault">
        /// Whether to cancel the default behavior.
        /// </param>
        private void OnVisualizeMenuClick(CommandBarButton ctrl, ref bool cancelDefault)
        {
            // Validate and create named Range.
            CreateNamedRange();
        }

        /// <summary>
        /// Handles Capture Viewpoint button click.
        /// </summary>
        /// <param name="sender">Capture Viewpoint button.</param>
        /// <param name="e">Ribbon Control Event arguments.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        private void OnCaptureViewpointButtonClick(object sender, RibbonControlEventArgs e)
        {
            
            try
            {
                Utility.IsWWTInstalled();
                WWTManager.IsValidMachine(Common.Globals.TargetMachine.MachineIP.ToString(), false);
                Perspective perspective = WWTManager.GetCameraView();

                if (perspective != null)
                {
                    var dialog = new CaptureViewpoint();
                    System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(dialog);
                    helper.Owner = (IntPtr)ThisAddIn.ExcelApplication.Hwnd;

                    perspective.Name = Resources.DefaultViewpointText;
                    var viewModel = new ViewpointViewModel(perspective);
                    viewModel.IsButtonEnabled = true;
                    viewModel.RequestClose += new EventHandler(dialog.OnRequestClose);
                    dialog.DataContext = viewModel;
                    dialog.ShowDialog();

                    // WPF dialog does not have dialog result
                    if (dialog.DialogResult.HasValue && dialog.DialogResult.Value)
                    {
                        if (!string.IsNullOrWhiteSpace(perspective.Name))
                        {
                            // Fire it only when valid value set
                            this.GetViewpointClicked.OnFire(perspective, new EventArgs());
                        }
                    }

                    viewModel.RequestClose -= new EventHandler(dialog.OnRequestClose);
                    dialog.Close();
                }
            }
            catch (CustomException ex)
            {
                Ribbon.ShowError(ex.HasCustomMessage ? ex.Message : Resources.LayerOperationError);
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
                Ribbon.ShowError(Resources.DefaultErrorMessage);
            }
        }

        /// <summary>
        /// Handles Go to Viewpoint button click.
        /// </summary>
        /// <param name="sender">Go to Viewpoint button.</param>
        /// <param name="e">Ribbon Control Event arguments.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        private void OnGotoViewpointButtonClick(object sender, RibbonControlEventArgs e)
        {
            RibbonGallery gallery = sender as RibbonGallery;
            if (gallery != null)
            {
                try
                {
                    Perspective viewpoint = (Perspective)gallery.SelectedItem.Tag;
                    if (viewpoint != null)
                    {
                        this.GotoViewpointClicked.OnFire(viewpoint, new EventArgs());
                    }
                }
                catch (Exception exception)
                {
                    Logger.LogException(exception);
                    Ribbon.ShowError(Resources.DefaultErrorMessage);
                }
            }
        }

        /// <summary>
        /// Handles Manage Viewpoint button click.
        /// </summary>
        /// <param name="sender">Manage Viewpoint button.</param>
        /// <param name="e">Ribbon Control Event arguments.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        private void OnManageViewpointButtonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                this.ManageViewpointClicked.OnFire(this, new EventArgs());
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
                Ribbon.ShowError(Resources.DefaultErrorMessage);
            }
        }

        /// <summary>
        /// Handles Go to Viewpoint From data button click.
        /// </summary>
        /// <param name="sender">Go to Viewpoint From data button.</param>
        /// <param name="e">Ribbon Control Event arguments.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        private void OnGotoViewpointFromDataButtonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                this.GotoViewpointFromDataClicked.OnFire(this, new EventArgs());
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
                Ribbon.ShowError(Resources.DefaultErrorMessage);
            }
        }

        /// <summary>
        /// Handles Target machine button click.
        /// </summary>
        /// <param name="sender">Target machine button.</param>
        /// <param name="e">Ribbon Control Event arguments.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        private void OnTargetMachineButtonClick(object sender, RibbonControlEventArgs e)
        {
         
            try
            {
                var dialog = new TargetMachinePane();
                System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(dialog);
                helper.Owner = (IntPtr)ThisAddIn.ExcelApplication.Hwnd;

                var viewModel = new TargetMachineViewModel(Common.Globals.TargetMachine.DisplayValue);
                viewModel.RequestClose += new EventHandler(dialog.OnRequestClose);

                dialog.DataContext = viewModel;

                dialog.ShowDialog();

                // WPF dialog does not have dialog result
                if (dialog.DialogResult.HasValue && dialog.DialogResult.Value)
                {
                    this.TargetMachineChanged.OnFire(Common.Globals.TargetMachine, new EventArgs());
                }

                viewModel.RequestClose -= new EventHandler(dialog.OnRequestClose);
                dialog.Close();
            }
            catch (CustomException ex)
            {
                Ribbon.ShowError(ex.HasCustomMessage ? ex.Message : Resources.DefaultErrorMessage);
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
                Ribbon.ShowError(Resources.DefaultErrorMessage);
            }
        }

        /// <summary>
        /// Handles Help button click.
        /// </summary>
        /// <param name="sender">Help button.</param>
        /// <param name="e">Ribbon Control Event arguments.</param>
        private void OnHelpButtonClick(object sender, RibbonControlEventArgs e)
        {
            StartProcess(Common.Constants.HelpLink);
        }

        /// <summary>
        /// Handles Contact Us button click.
        /// </summary>
        /// <param name="sender">Contact Us button.</param>
        /// <param name="e">Ribbon Control Event arguments.</param>
        private void OnContactUsButtonClick(object sender, RibbonControlEventArgs e)
        {
            StartProcess(Common.Constants.ContactUsLink);
        }

        /// <summary>
        /// Click event handler for the download updates button on the server
        /// </summary>
        /// <param name="sender">Download updates button</param>
        /// <param name="e">Ribbon control event args</param>
        private void OnDownloadUpdatesButtonClick(object sender, RibbonControlEventArgs e)
        {
            this.DownloadUpdatesButtonClicked.OnFire(sender, new EventArgs());
        }

        /// <summary>
        /// Handles Fetch climate button click.
        /// </summary>
        /// <param name="sender">Capture Viewpoint button.</param>
        /// <param name="e">Ribbon Control Event arguments.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        private void OnFetchClimateButtonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var dialog = new FetchClimateView();

                System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(dialog);
                helper.Owner = (IntPtr)ThisAddIn.ExcelApplication.Hwnd;

                var viewModel = new FetchClimateViewModel();
                viewModel.RequestClose += new EventHandler(dialog.OnRequestClose);
                dialog.DataContext = viewModel;
                dialog.ShowDialog();

                viewModel.RequestClose -= new EventHandler(dialog.OnRequestClose);
                dialog.Close();
            }
            catch (CustomException ex)
            {
                Ribbon.ShowError(ex.HasCustomMessage ? ex.Message : Resources.LayerOperationError);
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
                Ribbon.ShowError(Resources.DefaultErrorMessage);
            }
        }

        private void btnUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var dialog = new UpdateWizard();

                System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(dialog);
                helper.Owner = (IntPtr)ThisAddIn.ExcelApplication.Hwnd;

                var viewModel = new UpdateWizardViewModel();

                var latList = viewModel.ColumnHeaders.Where(lt => Common.Constants.LatSearchList.Contains(lt.ToLower()));
                var longList = viewModel.ColumnHeaders.Where(ln => Common.Constants.LonSearchList.Contains(ln.ToLower()));

                // if (string.Compare(viewModel.ColumnHeaders[0], "lat", true) == 0 && string.Compare(viewModel.ColumnHeaders[1], "lon", true) == 0)
                if (latList.Count() > 0 && longList.Count() > 0)
                {
                    viewModel.RequestClose += new EventHandler(dialog.OnRequestClose);
                    dialog.DataContext = viewModel;
                    dialog.ShowDialog();

                    viewModel.RequestClose -= new EventHandler(dialog.OnRequestClose);
                    dialog.Close();
                }
                else
                {
                    Ribbon.ShowError("The excel spreadsheet is not in expected format. Current worksheet should have \"lat\" and \"lon\" columns.");
                }
            }
            catch (CustomException ex)
            {
                Ribbon.ShowError(ex.HasCustomMessage ? ex.Message : Resources.LayerOperationError);
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
                Ribbon.ShowError(Resources.DefaultErrorMessage);
            }
        }

        private void btnAGUDemo_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Utility.IsWWTInstalled();
                WWTManager.IsValidMachine(Common.Globals.TargetMachine.MachineIP.ToString(), false);

                WorkflowController.Instance.ShowAGUDemo();
            }
            catch (CustomException ex)
            {
                Ribbon.ShowError(ex.HasCustomMessage ? ex.Message : Resources.LayerOperationError);
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
                Ribbon.ShowError(Resources.DefaultErrorMessage);
            }
        }

        private void btnIRISDemo_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var dialog = new Stations();

                System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(dialog);
                helper.Owner = (IntPtr)ThisAddIn.ExcelApplication.Hwnd;

                StationViewModel viewModel = new StationViewModel();

                viewModel.RequestClose += new EventHandler(dialog.OnRequestClose);
                dialog.DataContext = viewModel;
                dialog.ShowDialog();

                viewModel.RequestClose -= new EventHandler(dialog.OnRequestClose);
                dialog.Close();
            }
            catch (CustomException ex)
            {
                Ribbon.ShowError(ex.HasCustomMessage ? ex.Message : Resources.LayerOperationError);
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
                Ribbon.ShowError(Resources.DefaultErrorMessage);
            }
        }

        private void btnGetLocation_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Utility.IsWWTInstalled();
                WWTManager.IsValidMachine(Common.Globals.TargetMachine.MachineIP.ToString(), false);

                WorkflowController.Instance.UpdateCurrentLocation();
            }
            catch (CustomException ex)
            {
                Ribbon.ShowError(ex.HasCustomMessage ? ex.Message : Resources.LayerOperationError);
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
                Ribbon.ShowError(Resources.DefaultErrorMessage);
            }
        }

        #endregion Event handlers

        #region Private methods

        /// <summary>
        /// This function is used to validate the selection range and then invoke the VisualizeSelectionClicked event.
        /// which intern creates the named range in excel.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        private void CreateNamedRange()
        {
            try
            {
                Range selectedRange = ThisAddIn.ExcelApplication.Selection as Range;
                if (selectedRange != null && selectedRange.IsValid())
                {
                    VisualizeSelectionClicked.OnFire(this, new EventArgs());
                }
                else
                {
                    Ribbon.ShowError(Properties.Resources.InvalidSelectionRange);
                }
            }
            catch (CustomException exception)
            {
                Ribbon.ShowError(exception.HasCustomMessage ? exception.Message : Resources.DefaultErrorMessage);
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
                Ribbon.ShowError(Resources.DefaultErrorMessage);
            }
        }

        /// <summary>
        /// This function is used to add the context menu control for the worksheet.
        /// </summary>
        private void InitializeContextControls()
        {
            AddCommandBarButton("column");
            AddCommandBarButton("Cell");
            AddCommandBarButton("List Range Popup");
            AddCommandBarButton("PivotTable Context Menu");
        }

        /// <summary>
        /// Adds the "Visualize in WWT" button to the given Command bar of Excel.
        /// </summary>
        /// <param name="commandBarName">Command bar name</param>
        private void AddCommandBarButton(string commandBarName)
        {
            cellVisualizeMenu = (CommandBarButton)ThisAddIn.ExcelApplication.CommandBars[commandBarName].Controls.Add(MsoControlType.msoControlButton, Temporary: true);
            cellVisualizeMenu.Style = MsoButtonStyle.msoButtonCaption;
            cellVisualizeMenu.Caption = Properties.Resources.VisualizeMenuCaption;
            cellVisualizeMenu.Tag = Properties.Resources.VisualizeMenuTag;
            cellVisualizeMenu.Visible = true;
            cellVisualizeMenu.Click += OnVisualizeMenuClick;
        }

        #endregion Private methods

    }
}
