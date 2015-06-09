//-----------------------------------------------------------------------
// <copyright file="TargetMachineViewModel.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Windows.Input;
using Microsoft.Office.Interop.Excel;
using Microsoft.Research.Wwt.Excel.Addin.Properties;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Class represents the TargetMachine view model.
    /// </summary>
    public class TargetMachineViewModel : PropertyChangeBase
    {
        #region Private Properties

        private string name;
        private ICommand machineNameChangedCommand;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the TargetMachineViewModel class.
        /// </summary>
        /// <param name="machineName">
        /// target machine name
        /// </param>
        public TargetMachineViewModel(string machineName)
        {
            this.name = machineName;
            this.machineNameChangedCommand = new MachineNameChangeHandler(this);
        }

        #endregion

        #region CustomEvent

        /// <summary>
        /// Capture View window close request
        /// </summary>
        public event EventHandler RequestClose;

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the value of current perspective name
        /// </summary>
        public string Name
        {
            get
            {
                return this.name;
            }

            set
            {
                this.name = value;
            }
        }

        /// <summary>
        /// Gets the Machine name change command.
        /// </summary>
        public ICommand MachineNameChangedCommand
        {
            get { return this.machineNameChangedCommand; }
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Raise window close event
        /// </summary>
        public void OnRequestClose()
        {
            if (RequestClose != null)
            {
                RequestClose(this, EventArgs.Empty);
            }
        }

        #endregion

        #region Event Handler

        private class MachineNameChangeHandler : RelayCommand
        {
            private TargetMachineViewModel parent;

            public MachineNameChangeHandler(TargetMachineViewModel targetMachineViewModel)
            {
                this.parent = targetMachineViewModel;
            }

            public override void Execute(object parameter)
            {
                if (this.parent != null && parameter != null)
                {
                    string machineName = parameter as string;
                    try
                    {
                        ThisAddIn.ExcelApplication.Cursor = XlMousePointer.xlWait;
                        var machine = new TargetMachine(machineName.Trim());
                        Common.Globals.TargetMachine = machine;
                        ThisAddIn.ExcelApplication.Cursor = XlMousePointer.xlDefault;
                    }
                    catch (CustomException ex)
                    {
                        ThisAddIn.ExcelApplication.Cursor = XlMousePointer.xlDefault;
                        Ribbon.ShowError(ex.HasCustomMessage ? ex.Message : Resources.DefaultErrorMessage);
                        return;
                    }

                    // This is OK button click. So request window close
                    this.parent.OnRequestClose();
                }
            }
        }

        #endregion
    }
}
