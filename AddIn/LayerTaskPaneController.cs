using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using ExcelInterop = Microsoft.Office.Interop.Excel;
using AppEvents_Event = Microsoft.Office.Interop.Excel.AppEvents_Event;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Excel;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    internal delegate void LayerPaneChanged(LayerManagerPane pane);

    /// <summary>
    /// Excel major version numbers
    /// </summary>
    internal class ExcelVersion
    {
        public const string Excel2007 = "12.0";
        public const string Excel2010 = "14.0";
        public const string Excel2013 = "15.0";
    }

    /// <summary>
    /// 'Layer Manager TaskPane Controller' manages Office 2013 Single Document Interface (SDI).
    /// In 2013, each workbook has its own ribbon and taskpane, but multiple workbooks can be in the same process!
    /// It basically manages the interaction of WorkFlow controller with ribbon and taskpanes. It creates taskpanes for each workbook and
    /// associates the taskpane with workflow controller at runtime only for the activated instance of workbook, since workflow controller is only one instance.
    /// </summary>
    internal class LayerTaskPaneController
    {
        #region Data Members

        /// <summary>
        /// Event to notify that the current task pane has changed.
        /// Occurs when different workbook is activated or new workbook is created and then activated.
        /// </summary>
        public event LayerPaneChanged LayerPaneChangedEvent;

        /// <summary>
        /// reference to pane host associated with current the task pane
        /// </summary>
        private LayerManagerPaneHost currentPaneHost;

        /// <summary>
        /// reference to the current task pane
        /// </summary>
        private CustomTaskPane currentTaskPane;

        /// <summary>
        /// Host control for layer manager pane
        /// Association of workbook (workbook name) with Pane Host
        /// </summary>
        private Dictionary<string, LayerManagerPaneHost> workbookPaneHostMap;

        /// <summary>
        /// custom task pane instances
        /// Association of workbook (workbook name) with Task Pane
        /// </summary>
        private Dictionary<string, CustomTaskPane> workbookTaskPaneMap;

        /// <summary>
        /// Workbook name. Used when workbook is being renamed or saved.
        /// </summary>
        private string workbookName;

        /// <summary>
        /// String identifier to append to protect workbook names
        /// </summary>
        private const string protectedWorkbookIdentifier = "Protected_";

        #endregion
        
        #region Properties

        /// <summary>
        /// Gets the associated PaneHost of the activated Task Pane
        /// </summary>
        public LayerManagerPaneHost CurrentPaneHost
        {
            get { return currentPaneHost; }
            set { currentPaneHost = value; }
        }

        /// <summary>
        /// Gets the currently activated task Pane
        /// </summary>
        public CustomTaskPane CurrentTaskPane
        {
            get { return currentTaskPane; }
            set { currentTaskPane = value; }
        }

        /// <summary>
        /// Access to Workbook Pane Host Map
        /// </summary>
        public Dictionary<string, LayerManagerPaneHost> WorkbookPaneHostMap
        {
            get { return workbookPaneHostMap; }
        }

        /// <summary>
        /// Access to Workbook Task Pane Map
        /// </summary>
        public Dictionary<string, CustomTaskPane> WorkbookTaskPaneMap
        {
            get { return workbookTaskPaneMap; }
        }

        #endregion

        #region Constructor

        /// <summary>
            /// Singleton instance
            /// </summary>
        private static LayerTaskPaneController instance;

        /// <summary>
        /// Gets LayerTaskPaneManager instance
        /// </summary>
        public static LayerTaskPaneController Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new LayerTaskPaneController();
                }
                return instance;
            }
        }

        /// <summary>
        /// Private constructor
        /// </summary>
        LayerTaskPaneController()
        {
            workbookPaneHostMap = new Dictionary<string, LayerManagerPaneHost>();
            workbookTaskPaneMap = new Dictionary<string, CustomTaskPane>();
        }

        #endregion

        #region Event Handlers

        /// <summary>
        /// New workbook event
        /// </summary>
        void OnExcelNewWorkbook(ExcelInterop.Workbook Wb)
        {
            //Each time a new workbook is created, create a task pane for it. 
            //This event is followed by, workbook activated event
            AddNewTaskPane(Wb.Name);
        }

        /// <summary>
        /// Open workbook event
        /// </summary>
        void OnExcelWorkbookOpen(ExcelInterop.Workbook Wb)
        {
            AddNewTaskPane(Wb.Name);

            //When CSV files are opened. Events are fired in wrong order!
            //Workbook Activate is caled before Workbook Open. We require Activate to be called *after* Open
            //TODO: Need to investigate why this is so
            if (Wb.FileFormat == ExcelInterop.XlFileFormat.xlCSV)
                OnExcelWorkbookActivate(Wb);
        }

        /// <summary>
        /// Open protected workbook event
        /// </summary>
        void OnProtectedExcelWorkbookOpen(ExcelInterop.ProtectedViewWindow Pvw)
        {
            AddNewTaskPane(protectedWorkbookIdentifier + Pvw.Workbook.Name);

            //When CSV files are opened. Events are fired in wrong order!
            //Workbook Activate is caled before Workbook Open. We require Activate to be called *after* Open
            //TODO: Need to investigate why this is so
            if (Pvw.Workbook.FileFormat == ExcelInterop.XlFileFormat.xlCSV)
                OnProtectedExcelWorkbookActivate(Pvw);
        }

        /// <summary>
        /// Updates this controller to refer to the taskPane & paneHost of the activated workbook
        /// And then, notify that the current activated taskpane is changed
        /// </summary>
        void OnExcelWorkbookActivate(ExcelInterop.Workbook Wb)
        {
            //Workaround for, excel in Protected mode.
            //When in protected mode (excel launches two processes), and when user enables editing it looks excel closes all taskpanes. 
            //Window handlle of Task pane created by us is no longer valid. We catch this exception and attempt to create new task pane
            try
            {
                if(LayerTaskPaneController.Instance.CurrentTaskPane != null)
                    LayerTaskPaneController.Instance.CurrentTaskPane.Visible = LayerTaskPaneController.Instance.CurrentTaskPane.Visible;
            }
            //Workaround for, excel in Protected mode.
            //When in protected mode (excel launches two processes), and when user enables editing it looks excel closes all taskpanes. 
            //Window handlle of Task pane created by us is no longer valid. We catch this exception and attempt to create new task pane
            catch (System.Runtime.InteropServices.COMException exception)
            {
                if (exception.ErrorCode == -2146827864)
                {
                    //Remove this invalid task pane
                    RemoveTaskPane(Wb.Name);

                    //Add new task pane
                    AddNewTaskPane(Wb.Name);
                }
            }

            //Set the respective taskpane for the activated workbook
            SetCurrentTaskPane(Wb.Name);
        }

        /// <summary>
        /// On protected excel workbook activate
        /// </summary>
        void OnProtectedExcelWorkbookActivate(ExcelInterop.ProtectedViewWindow Pvw)
        {
            SetCurrentTaskPane(protectedWorkbookIdentifier + Pvw.Workbook.Name);
        }

        /// <summary>
        /// Updates controller to refer to the taskPane & paneHost of the activated workbook
        /// </summary>
        void OnExcelWorkbookBeforeClose(ExcelInterop.Workbook Wb, ref bool Cancel)
        {
            RemoveTaskPane(Wb.Name);
        }

        /// <summary>
        /// 
        /// </summary>
        void OnProtectedExcelWorkbookBeforeClose(ExcelInterop.ProtectedViewWindow Pvw, ExcelInterop.XlProtectedViewCloseReason Reason, ref bool Cancel)
        {
            RemoveTaskPane(protectedWorkbookIdentifier + Pvw.Workbook.Name);
        }

        /// <summary>
        /// Before workbook save event
        /// </summary>
        void OnExcelWorkbookBeforeSave(ExcelInterop.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            workbookName = Wb.Name;
        }

        /// <summary>
        /// After Workbook save event
        /// Update map keys if workbook name has changed
        /// </summary>
        void OnExcelWorkbookAfterSave(ExcelInterop.Workbook Wb, bool Success)
        {
            string oldBookName = LayerTaskPaneController.Instance.workbookName;

            // When workbook name changes, the corresponding maps value is out of sync with its key (workbook name). Need to sync
            if(Wb.Name.Equals(oldBookName) == false)
                UpdateMapKeys(Wb);
        }

        /// <summary>
        /// Handles CustomTaskPane Visibility
        /// </summary>
        /// <param name="sender">Custom Task Pane</param>
        /// <param name="e">Event arguments</param>
        void OnCustomTaskPaneVisibleChanged(object sender, EventArgs e)
        {
            Microsoft.Office.Tools.CustomTaskPane custTaskPane = (Microsoft.Office.Tools.CustomTaskPane)sender;
            Globals.Ribbons.Ribbon.layerManagerButton.Checked = custTaskPane.Visible;
            if (Globals.Ribbons.Ribbon.layerManagerButton.Checked)
            {
              
            }
            else
            {
              
            }
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Register for excel events. Core of functionality lies on these fired events.
        /// When events are fired, new task panes are created, taskpanes are activated, made visible, ect.
        /// </summary>
        void RegisterforExcelEvents()
        {
            ((ExcelInterop.AppEvents_Event)Globals.ThisAddIn.Application).NewWorkbook += 
                new ExcelInterop.AppEvents_NewWorkbookEventHandler(OnExcelNewWorkbook);

            ((ExcelInterop.AppEvents_Event)Globals.ThisAddIn.Application).WorkbookOpen += 
                new ExcelInterop.AppEvents_WorkbookOpenEventHandler(OnExcelWorkbookOpen);

            ((ExcelInterop.AppEvents_Event)Globals.ThisAddIn.Application).ProtectedViewWindowOpen += 
                new ExcelInterop.AppEvents_ProtectedViewWindowOpenEventHandler(OnProtectedExcelWorkbookOpen);

            ((ExcelInterop.AppEvents_Event)Globals.ThisAddIn.Application).WorkbookActivate += 
                new ExcelInterop.AppEvents_WorkbookActivateEventHandler(OnExcelWorkbookActivate);

            ((ExcelInterop.AppEvents_Event)Globals.ThisAddIn.Application).ProtectedViewWindowActivate +=
                new ExcelInterop.AppEvents_ProtectedViewWindowActivateEventHandler(OnProtectedExcelWorkbookActivate);

            ((ExcelInterop.AppEvents_Event)Globals.ThisAddIn.Application).WorkbookBeforeClose += 
                new ExcelInterop.AppEvents_WorkbookBeforeCloseEventHandler(OnExcelWorkbookBeforeClose);

            ((ExcelInterop.AppEvents_Event)Globals.ThisAddIn.Application).ProtectedViewWindowBeforeClose +=
                new ExcelInterop.AppEvents_ProtectedViewWindowBeforeCloseEventHandler(OnProtectedExcelWorkbookBeforeClose);

            ((ExcelInterop.AppEvents_Event)Globals.ThisAddIn.Application).WorkbookBeforeSave += 
                new ExcelInterop.AppEvents_WorkbookBeforeSaveEventHandler(OnExcelWorkbookBeforeSave);

            ((ExcelInterop.AppEvents_Event)Globals.ThisAddIn.Application).WorkbookAfterSave += 
                new ExcelInterop.AppEvents_WorkbookAfterSaveEventHandler(OnExcelWorkbookAfterSave);
        }

        /// <summary>
        /// Since workbook name has changed, update map keys to reflect with the new name as key
        /// </summary>
        /// <param name="Wb"></param>
        void UpdateMapKeys(ExcelInterop.Workbook Wb)
        {
            string oldName = LayerTaskPaneController.Instance.workbookName;
            
            LayerManagerPaneHost paneHost;
            if(LayerTaskPaneController.Instance.workbookPaneHostMap.ContainsKey(oldName))
            {
                //get reference to pane host corresponding to the old key
                paneHost = LayerTaskPaneController.Instance.workbookPaneHostMap[oldName];

                //remove the old key-value
                LayerTaskPaneController.Instance.workbookPaneHostMap.Remove(oldName);

                //add the updated key-value
                LayerTaskPaneController.Instance.workbookPaneHostMap.Add(Wb.Name, paneHost);
            }

            CustomTaskPane taskPane;
            if (LayerTaskPaneController.Instance.workbookTaskPaneMap.ContainsKey(oldName))
            {
                //get reference to task pane corresponding to the old key
                taskPane = LayerTaskPaneController.Instance.workbookTaskPaneMap[oldName];
                
                //remove the old key-value
                LayerTaskPaneController.Instance.workbookTaskPaneMap.Remove(oldName);

                //add the updated key-value
                LayerTaskPaneController.Instance.workbookTaskPaneMap.Add(Wb.Name, taskPane);
            }

            LayerTaskPaneController.Instance.workbookName = string.Empty;
        }

        /// <summary>
        /// Adds a new item to workbookPaneHostMap dictionary
        /// </summary>
        /// <param name="WorkbookName"></param>
        /// <param name="layerManagerPaneHost"></param>
        void AddToWorkbookPaneHostMap(string WorkbookName, LayerManagerPaneHost layerManagerPaneHost)
        {
            //TODO: Throw exception
            //if (workbookPaneHostMap.ContainsKey(WorkbookName))

            workbookPaneHostMap.Add(WorkbookName, layerManagerPaneHost);
        }

        /// <summary>
        /// Adds a new item to workbookTaskPaneMap dictionary
        /// </summary>
        /// <param name="WorkbookName"></param>
        /// <param name="customTaskPane"></param>
        void AddToWorkbookTaskPaneMap(string WorkbookName, CustomTaskPane customTaskPane)
        {
            //TODO: Throw exception
            //if (workbookTaskPaneMap.ContainsKey(WorkbookName))

            workbookTaskPaneMap.Add(WorkbookName, customTaskPane);
        }

        /// <summary>
        /// Create a task pane and Also Maintain association of workbook with taskpane.
        /// </summary>
        /// <param name="WorkbookName">Name of workbook for which the task pane is being created.</param>
        void AddNewTaskPane(string WorkbookName)
        {
            LayerTaskPaneController.Instance.CurrentPaneHost = new LayerManagerPaneHost();
            LayerTaskPaneController.Instance.CurrentTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(LayerTaskPaneController.Instance.CurrentPaneHost, Properties.Resources.LayerManagerPane);

            LayerTaskPaneController.Instance.CurrentTaskPane.DockPosition = Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
            LayerTaskPaneController.Instance.CurrentTaskPane.VisibleChanged += new EventHandler(this.OnCustomTaskPaneVisibleChanged);
            LayerTaskPaneController.Instance.CurrentTaskPane.Visible = false;
            LayerTaskPaneController.Instance.CurrentTaskPane.Width = 315;

            //In case of MDI, only one task pane is created & map collection is never accessed
            if (LayerTaskPaneController.IsExcelInstanceSDI)
            {
                LayerTaskPaneController.Instance.AddToWorkbookPaneHostMap(WorkbookName, LayerTaskPaneController.Instance.CurrentPaneHost);
                LayerTaskPaneController.Instance.AddToWorkbookTaskPaneMap(WorkbookName, LayerTaskPaneController.Instance.CurrentTaskPane);
            }
        }

        /// <summary>
        /// Sets the current taskpane and fires event notifying change of current task pane
        /// </summary>
        /// <param name="WorkbookName">Name of the Workbook that is associated with the taskpane to be set</param>
        void SetCurrentTaskPane(string WorkbookName)
        {
            if (LayerTaskPaneController.Instance.WorkbookPaneHostMap.ContainsKey(WorkbookName))
            {
                LayerTaskPaneController.Instance.CurrentPaneHost = LayerTaskPaneController.Instance.WorkbookPaneHostMap[WorkbookName];
                if (LayerTaskPaneController.Instance.WorkbookTaskPaneMap.ContainsKey(WorkbookName))
                {
                    LayerTaskPaneController.Instance.CurrentTaskPane = LayerTaskPaneController.Instance.WorkbookTaskPaneMap[WorkbookName];
                }

                //Notify that the current task pane has changed
                if(LayerPaneChangedEvent != null)
                    LayerPaneChangedEvent(LayerTaskPaneController.Instance.CurrentPaneHost.LayerManagerPane);
            }
        }

        /// <summary>
        /// Removes the taskpane associated with the specified workbookName
        /// </summary>
        void RemoveTaskPane(string WorkbookName)
        {
            if (LayerTaskPaneController.Instance.WorkbookPaneHostMap.ContainsKey(WorkbookName) ||
                LayerTaskPaneController.Instance.WorkbookTaskPaneMap.ContainsKey(WorkbookName))
            {
                CustomTaskPane paneToRemove = null;

                //If background excel is closed, then CurrentTaskPane is *not* the one that is being removed. 
                //We need to remove the objects corresponding to what is being removed. 
                if (LayerTaskPaneController.Instance.CurrentTaskPane != LayerTaskPaneController.Instance.WorkbookTaskPaneMap[WorkbookName])
                {
                    paneToRemove = LayerTaskPaneController.Instance.WorkbookTaskPaneMap[WorkbookName];
                }
                else   //activated excel workbook is the excel workbook which is being removed
                {
                    paneToRemove = LayerTaskPaneController.Instance.CurrentTaskPane;

                    LayerTaskPaneController.Instance.CurrentPaneHost = null;
                    LayerTaskPaneController.Instance.CurrentTaskPane = null;
                }

                try
                {
                    paneToRemove.Visible = false;
                }
                catch (Exception)
                {
                }

                Globals.ThisAddIn.CustomTaskPanes.Remove(paneToRemove);

                LayerTaskPaneController.Instance.WorkbookPaneHostMap.Remove(WorkbookName);
                LayerTaskPaneController.Instance.WorkbookTaskPaneMap.Remove(WorkbookName);

                //NOTE: If activated excel workbook was closed, WorkflowController will be in invalid state since the reference object 
                //(WorkflowController.layerManagerPaneInstance) no longer exists (CurrentTaskPane). 
                //But after on close, next workbook activate event will get triggered, where it is updated to refer to a valid pane.
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Checks if Excel is SDI (Office 2013) or MDI (Office 2010 / 2007). 
        /// </summary>
        public static bool IsExcelInstanceSDI
        {
            get
            {
                switch (Globals.ThisAddIn.Application.Version)
                {
                    case ExcelVersion.Excel2007:
                        return false;

                    case ExcelVersion.Excel2010:
                        return false;

                    case ExcelVersion.Excel2013:
                        return true;

                    //Other versions not supported
                    default:
                        throw new NotImplementedException("This version of Excel is not supported!");
                }
            }
        }

        /// <summary>
        /// Initializes the controller.
        /// Registers for excel events based on SDI / MDI and creates task pane appropriately
        /// </summary>
        public void Initialize()
        {
            if (LayerTaskPaneController.IsExcelInstanceSDI)
            {
                //Register for excel events. 
                //When new workbook event is fired, we create a task pane for that workbook and maintain their association
                RegisterforExcelEvents();

                //when excel is launched from installer just after installing, then a workbook is already opened and addin is yet to be loaded. 
                //Hence excel Events are not caught and no item has been added to task pane map
                //Explicitly add to the task pane map & set the current task pane
                //TODO: Installer to prevent installation if is excel is opened!
                if (Globals.ThisAddIn.Application.Workbooks.Count == 1 && LayerTaskPaneController.Instance.WorkbookTaskPaneMap.Count == 0)
                {
                    if (Globals.ThisAddIn.Application.ActiveWorkbook != null)
                    {
                        AddNewTaskPane(Globals.ThisAddIn.Application.ActiveWorkbook.Name);
                        SetCurrentTaskPane(Globals.ThisAddIn.Application.ActiveWorkbook.Name);
                    }
                }
            }
            else
            {
                //Only one task pane will be created for excel when in MDI. (Initialize workbook name to any dummy value.)
                string wbName = "MDITaskPane";
                AddNewTaskPane(wbName);
            }
        }

        #endregion
    }
}

//TODO
//Disable ViewInWWT on all workbooks of same process

//TOCHECK
//Is ribbon in sync with workflowcontroller?
//model update check - workflowcontroller.UpdateLayerManagerPaneInstance()
//protected workbook should disable ribbon buttons
