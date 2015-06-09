//-----------------------------------------------------------------------
// <copyright file="Ribbon.Designer.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using System.Diagnostics.CodeAnalysis;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        internal Microsoft.Office.Tools.Ribbon.RibbonTab wwteRibbonTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup visualizeGroup;
        internal Office.Tools.Ribbon.RibbonButton visualizeSelectionButton;
        internal Office.Tools.Ribbon.RibbonToggleButton layerManagerButton;
        internal Office.Tools.Ribbon.RibbonGroup viewpointsGroup;

        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Constructor
        /// </summary>
        /// Do not add ExcludeFromCodeCoverage attribute at class class for Ribbon. This is is a partial class and
        /// there are custom implementations in Ribbon.cs which needs to be covered under code coverage.
        [ExcludeFromCodeCoverage]
        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        [ExcludeFromCodeCoverage]
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        [ExcludeFromCodeCoverage]
        private void InitializeComponent()
        {
            this.wwteRibbonTab = this.Factory.CreateRibbonTab();
            this.dummyGroup = this.Factory.CreateRibbonGroup();
            this.visualizeGroup = this.Factory.CreateRibbonGroup();
            this.visualizeSelectionButton = this.Factory.CreateRibbonButton();
            this.tgBtnAutoMove = this.Factory.CreateRibbonToggleButton();
            this.layerManagerButton = this.Factory.CreateRibbonToggleButton();
            this.viewSamplesGallery = this.Factory.CreateRibbonGallery();
            this.viewpointsGroup = this.Factory.CreateRibbonGroup();
            this.captureViewpoint = this.Factory.CreateRibbonButton();
            this.gotoViewpointGallery = this.Factory.CreateRibbonGallery();
            this.manageViewpoint = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.gotoViewpointFromData = this.Factory.CreateRibbonButton();
            this.machineNameGroup = this.Factory.CreateRibbonGroup();
            this.targetMachineButton = this.Factory.CreateRibbonButton();
            this.helpGroup = this.Factory.CreateRibbonGroup();
            this.contactusButton = this.Factory.CreateRibbonButton();
            this.feedbackGallery = this.Factory.CreateRibbonGallery();
            this.helpButton = this.Factory.CreateRibbonButton();
            this.updateGroup = this.Factory.CreateRibbonGroup();
            this.downloadUpdatesButton = this.Factory.CreateRibbonButton();
            this.fetchClimateGroup = this.Factory.CreateRibbonGroup();
            this.btnUpdate = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.fetchClimateButton = this.Factory.CreateRibbonButton();
            this.btnAGUDemo = this.Factory.CreateRibbonButton();
            this.btnIRISDemo = this.Factory.CreateRibbonButton();
            this.btnAboutServices = this.Factory.CreateRibbonButton();
            this.btnGetLocation = this.Factory.CreateRibbonButton();
            this.wwteRibbonTab.SuspendLayout();
            this.visualizeGroup.SuspendLayout();
            this.viewpointsGroup.SuspendLayout();
            this.machineNameGroup.SuspendLayout();
            this.helpGroup.SuspendLayout();
            this.updateGroup.SuspendLayout();
            this.fetchClimateGroup.SuspendLayout();
            // 
            // wwteRibbonTab
            // 
            this.wwteRibbonTab.Groups.Add(this.dummyGroup);
            this.wwteRibbonTab.Groups.Add(this.visualizeGroup);
            this.wwteRibbonTab.Groups.Add(this.viewpointsGroup);
            this.wwteRibbonTab.Groups.Add(this.machineNameGroup);
            this.wwteRibbonTab.Groups.Add(this.helpGroup);
            this.wwteRibbonTab.Groups.Add(this.updateGroup);
            this.wwteRibbonTab.Groups.Add(this.fetchClimateGroup);
            this.wwteRibbonTab.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.RibbonTabName;
            this.wwteRibbonTab.Name = "wwteRibbonTab";
            // 
            // dummyGroup
            // 
            this.dummyGroup.Label = "dummyGroup";
            this.dummyGroup.Name = "dummyGroup";
            this.dummyGroup.Visible = false;
            // 
            // visualizeGroup
            // 
            this.visualizeGroup.Items.Add(this.visualizeSelectionButton);
            this.visualizeGroup.Items.Add(this.tgBtnAutoMove);
            this.visualizeGroup.Items.Add(this.layerManagerButton);
            this.visualizeGroup.Items.Add(this.viewSamplesGallery);
            this.visualizeGroup.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.RibbonVisualizeGroupName;
            this.visualizeGroup.Name = "visualizeGroup";
            // 
            // visualizeSelectionButton
            // 
            this.visualizeSelectionButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.visualizeSelectionButton.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.VisualizeSelection;
            this.visualizeSelectionButton.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.RibbonVisualizeSelectionButtonName;
            this.visualizeSelectionButton.Name = "visualizeSelectionButton";
            this.visualizeSelectionButton.ScreenTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.VisualizeSelectionButtonScreenTip;
            this.visualizeSelectionButton.ShowImage = true;
            this.visualizeSelectionButton.SuperTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.VisualizeSelectionButtonToolTip;
            this.visualizeSelectionButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnVisualizeSelectionButtonClick);
            // 
            // tgBtnAutoMove
            // 
            this.tgBtnAutoMove.Label = "Auto Move";
            this.tgBtnAutoMove.Name = "tgBtnAutoMove";
            // 
            // layerManagerButton
            // 
            this.layerManagerButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.layerManagerButton.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.LayerManager;
            this.layerManagerButton.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.RibbonLayerManagerButtonName;
            this.layerManagerButton.Name = "layerManagerButton";
            this.layerManagerButton.ScreenTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.LayerManagerButtonScreenTip;
            this.layerManagerButton.ShowImage = true;
            this.layerManagerButton.SuperTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.LayerManagerButtonToolTip;
            this.layerManagerButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnLayerButtonClick);
            // 
            // viewSamplesGallery
            // 
            this.viewSamplesGallery.ColumnCount = 1;
            this.viewSamplesGallery.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.viewSamplesGallery.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.ViewSamples;
            this.viewSamplesGallery.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.ViewSamplesGalleryLabel;
            this.viewSamplesGallery.Name = "viewSamplesGallery";
            this.viewSamplesGallery.ScreenTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.ViewSamplesButtonScreenTip;
            this.viewSamplesGallery.ShowImage = true;
            this.viewSamplesGallery.SuperTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.ViewSamplesButtonToolTip;
            this.viewSamplesGallery.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnViewSamplesGalleryClick);
            // 
            // viewpointsGroup
            // 
            this.viewpointsGroup.Items.Add(this.captureViewpoint);
            this.viewpointsGroup.Items.Add(this.gotoViewpointGallery);
            this.viewpointsGroup.Items.Add(this.manageViewpoint);
            this.viewpointsGroup.Items.Add(this.separator1);
            this.viewpointsGroup.Items.Add(this.gotoViewpointFromData);
            this.viewpointsGroup.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.ViewpointsGroupLabel;
            this.viewpointsGroup.Name = "viewpointsGroup";
            // 
            // captureViewpoint
            // 
            this.captureViewpoint.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.captureViewpoint.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.CaptureViewpoint;
            this.captureViewpoint.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.CaptureViewpointText;
            this.captureViewpoint.Name = "captureViewpoint";
            this.captureViewpoint.ScreenTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.CaptureViewpointButtonScreenTip;
            this.captureViewpoint.ShowImage = true;
            this.captureViewpoint.SuperTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.CaptureViewpointButtonSuperTip;
            this.captureViewpoint.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnCaptureViewpointButtonClick);
            // 
            // gotoViewpointGallery
            // 
            this.gotoViewpointGallery.ColumnCount = 1;
            this.gotoViewpointGallery.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.gotoViewpointGallery.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.GotoViewpoint;
            this.gotoViewpointGallery.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.GotoViewpointButtonLabel;
            this.gotoViewpointGallery.Name = "gotoViewpointGallery";
            this.gotoViewpointGallery.ScreenTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.GotoViewpointButtonScreenTip;
            this.gotoViewpointGallery.ShowImage = true;
            this.gotoViewpointGallery.SuperTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.GotoViewpointButtonSuperTip;
            this.gotoViewpointGallery.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnGotoViewpointButtonClick);
            // 
            // manageViewpoint
            // 
            this.manageViewpoint.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.manageViewpoint.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.ManageViewpoint;
            this.manageViewpoint.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.ManageViewpointButtonLabel;
            this.manageViewpoint.Name = "manageViewpoint";
            this.manageViewpoint.ScreenTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.ManageViewpointButtonScreenTip;
            this.manageViewpoint.ShowImage = true;
            this.manageViewpoint.SuperTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.ManageViewpointButtonSuperTip;
            this.manageViewpoint.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnManageViewpointButtonClick);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // gotoViewpointFromData
            // 
            this.gotoViewpointFromData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.gotoViewpointFromData.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.ViewpointFromData;
            this.gotoViewpointFromData.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.GotoViewpointFromDataButtonLabel;
            this.gotoViewpointFromData.Name = "gotoViewpointFromData";
            this.gotoViewpointFromData.ScreenTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.GotoViewpointFromDataButtonScreenTip;
            this.gotoViewpointFromData.ShowImage = true;
            this.gotoViewpointFromData.SuperTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.GotoViewpointFromDataButtonSuperTip;
            this.gotoViewpointFromData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnGotoViewpointFromDataButtonClick);
            // 
            // machineNameGroup
            // 
            this.machineNameGroup.Items.Add(this.targetMachineButton);
            this.machineNameGroup.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.TargetMachineLabel;
            this.machineNameGroup.Name = "machineNameGroup";
            // 
            // targetMachineButton
            // 
            this.targetMachineButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.targetMachineButton.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.Network;
            this.targetMachineButton.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.TargetMachineButtonLabel;
            this.targetMachineButton.Name = "targetMachineButton";
            this.targetMachineButton.ScreenTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.TargetMachineButtonScreenTip;
            this.targetMachineButton.ShowImage = true;
            this.targetMachineButton.SuperTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.TargetMachineButtonToolTip;
            this.targetMachineButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnTargetMachineButtonClick);
            // 
            // helpGroup
            // 
            this.helpGroup.Items.Add(this.contactusButton);
            this.helpGroup.Items.Add(this.feedbackGallery);
            this.helpGroup.Items.Add(this.helpButton);
            this.helpGroup.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.HelpGroupLabel;
            this.helpGroup.Name = "helpGroup";
            // 
            // contactusButton
            // 
            this.contactusButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.contactusButton.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.Mail;
            this.contactusButton.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.ContactUsButtonLabel;
            this.contactusButton.Name = "contactusButton";
            this.contactusButton.ScreenTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.ContactUsButtonScreenTip;
            this.contactusButton.ShowImage = true;
            this.contactusButton.SuperTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.ContactUsButtonToolTip;
            this.contactusButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnContactUsButtonClick);
            // 
            // feedbackGallery
            // 
            this.feedbackGallery.ColumnCount = 1;
            this.feedbackGallery.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.feedbackGallery.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.Feedback;
            this.feedbackGallery.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.FeedbackGalleryLabel;
            this.feedbackGallery.Name = "feedbackGallery";
            this.feedbackGallery.ScreenTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.FeedbackGalleryScreenTip;
            this.feedbackGallery.ShowImage = true;
            this.feedbackGallery.SuperTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.FeedbackGallerySuperTip;
            this.feedbackGallery.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnFeedbackGalleryClick);
            // 
            // helpButton
            // 
            this.helpButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.helpButton.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.Help;
            this.helpButton.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.HelpButtonLabel;
            this.helpButton.Name = "helpButton";
            this.helpButton.ScreenTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.HelpButtonScreenTip;
            this.helpButton.ShowImage = true;
            this.helpButton.SuperTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.HelpButtonToolTip;
            this.helpButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnHelpButtonClick);
            // 
            // updateGroup
            // 
            this.updateGroup.Items.Add(this.downloadUpdatesButton);
            this.updateGroup.Label = "Update";
            this.updateGroup.Name = "updateGroup";
            this.updateGroup.Visible = false;
            // 
            // downloadUpdatesButton
            // 
            this.downloadUpdatesButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.downloadUpdatesButton.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.DownloadUpdate;
            this.downloadUpdatesButton.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.DownloadUpdatesButtonLabel;
            this.downloadUpdatesButton.Name = "downloadUpdatesButton";
            this.downloadUpdatesButton.ScreenTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.DownloadUpdatesButtonScreenTip;
            this.downloadUpdatesButton.ShowImage = true;
            this.downloadUpdatesButton.SuperTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.DownloadUpdatesButtonToolTip;
            this.downloadUpdatesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnDownloadUpdatesButtonClick);
            // 
            // fetchClimateGroup
            // 
            this.fetchClimateGroup.Items.Add(this.btnGetLocation);
            this.fetchClimateGroup.Items.Add(this.btnUpdate);
            this.fetchClimateGroup.Items.Add(this.separator2);
            this.fetchClimateGroup.Items.Add(this.fetchClimateButton);
            this.fetchClimateGroup.Items.Add(this.btnAGUDemo);
            this.fetchClimateGroup.Items.Add(this.btnIRISDemo);
            this.fetchClimateGroup.Items.Add(this.btnAboutServices);
            this.fetchClimateGroup.Label = "Services";
            this.fetchClimateGroup.Name = "fetchClimateGroup";
            // 
            // btnUpdate
            // 
            this.btnUpdate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdate.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.GenerateColumns;
            this.btnUpdate.Label = "Cuboid";
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.ShowImage = true;
            this.btnUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdate_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // fetchClimateButton
            // 
            this.fetchClimateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.fetchClimateButton.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.Sky;
            this.fetchClimateButton.Label = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.FetchClimateViewText;
            this.fetchClimateButton.Name = "fetchClimateButton";
            this.fetchClimateButton.ScreenTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.FetchClimateScreenTip;
            this.fetchClimateButton.ShowImage = true;
            this.fetchClimateButton.SuperTip = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.FetchClimateToolTip;
            this.fetchClimateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnFetchClimateButtonClick);
            // 
            // btnAGUDemo
            // 
            this.btnAGUDemo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAGUDemo.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.demoicon;
            this.btnAGUDemo.Label = "FC Demo";
            this.btnAGUDemo.Name = "btnAGUDemo";
            this.btnAGUDemo.ShowImage = true;
            this.btnAGUDemo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAGUDemo_Click);
            // 
            // btnIRISDemo
            // 
            this.btnIRISDemo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnIRISDemo.Label = "IRIS";
            this.btnIRISDemo.Name = "btnIRISDemo";
            this.btnIRISDemo.ShowImage = true;
            this.btnIRISDemo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnIRISDemo_Click);
            // 
            // btnAboutServices
            // 
            this.btnAboutServices.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAboutServices.Description = "Proof Of Concept explorations in Excel as a data service consumer";
            this.btnAboutServices.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.MapColumnInfo;
            this.btnAboutServices.Label = "About Services";
            this.btnAboutServices.Name = "btnAboutServices";
            this.btnAboutServices.ScreenTip = "Proof Of Concept explorations in Excel as a data service consumer";
            this.btnAboutServices.ShowImage = true;
            this.btnAboutServices.SuperTip = "Proof Of Concept explorations in Excel as a data service consumer";
            // 
            // btnGetLocation
            // 
            this.btnGetLocation.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGetLocation.Image = global::Microsoft.Research.Wwt.Excel.Addin.Properties.Resources.location;
            this.btnGetLocation.Label = "Get Location";
            this.btnGetLocation.Name = "btnGetLocation";
            this.btnGetLocation.ShowImage = true;
            this.btnGetLocation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetLocation_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.wwteRibbonTab);
            this.Close += new System.EventHandler(this.OnRibbonClose);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.OnRibbonLoad);
            this.wwteRibbonTab.ResumeLayout(false);
            this.wwteRibbonTab.PerformLayout();
            this.visualizeGroup.ResumeLayout(false);
            this.visualizeGroup.PerformLayout();
            this.viewpointsGroup.ResumeLayout(false);
            this.viewpointsGroup.PerformLayout();
            this.machineNameGroup.ResumeLayout(false);
            this.machineNameGroup.PerformLayout();
            this.helpGroup.ResumeLayout(false);
            this.helpGroup.PerformLayout();
            this.updateGroup.ResumeLayout(false);
            this.updateGroup.PerformLayout();
            this.fetchClimateGroup.ResumeLayout(false);
            this.fetchClimateGroup.PerformLayout();

        }


        #endregion

        internal Office.Tools.Ribbon.RibbonGallery viewSamplesGallery;
        internal Office.Tools.Ribbon.RibbonButton targetMachineButton;
        internal Office.Tools.Ribbon.RibbonGroup machineNameGroup;
        internal Office.Tools.Ribbon.RibbonGroup dummyGroup;
        internal Office.Tools.Ribbon.RibbonButton captureViewpoint;
        internal Office.Tools.Ribbon.RibbonButton manageViewpoint;
        internal Office.Tools.Ribbon.RibbonButton gotoViewpointFromData;
        internal Office.Tools.Ribbon.RibbonGallery gotoViewpointGallery;
        internal Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Office.Tools.Ribbon.RibbonGroup helpGroup;
        internal Office.Tools.Ribbon.RibbonButton helpButton;
        internal Office.Tools.Ribbon.RibbonButton contactusButton;
        internal Office.Tools.Ribbon.RibbonGallery feedbackGallery;
        internal Office.Tools.Ribbon.RibbonButton downloadUpdatesButton;
        internal System.ComponentModel.BackgroundWorker checkUpdatesWorker;
        internal System.ComponentModel.BackgroundWorker downloadLinkWorker;
        internal Office.Tools.Ribbon.RibbonGroup updateGroup;
        internal Office.Tools.Ribbon.RibbonGroup fetchClimateGroup;
        internal Office.Tools.Ribbon.RibbonButton fetchClimateButton;
        internal Office.Tools.Ribbon.RibbonButton btnUpdate;
        internal Office.Tools.Ribbon.RibbonButton btnAGUDemo;
        internal Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Office.Tools.Ribbon.RibbonButton btnAboutServices;
        internal Office.Tools.Ribbon.RibbonToggleButton tgBtnAutoMove;
        internal Office.Tools.Ribbon.RibbonButton btnIRISDemo;
        internal Office.Tools.Ribbon.RibbonButton btnGetLocation;
    }

    [ExcludeFromCodeCoverage]
    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}

