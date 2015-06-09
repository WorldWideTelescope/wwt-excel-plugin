// -----------------------------------------------------------------------
// <copyright file="UpdateWizardViewModel.cs">
//  View model for fetch climate view.
// </copyright>
// -----------------------------------------------------------------------

namespace Microsoft.Research.Wwt.Excel.Addin
{
    using System;
    using System.Linq;
    using System.Collections.Generic;
    using System.Windows.Input;
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Research.Wwt.Excel.Common;
    using Microsoft.Research.Wwt.Excel.Addin.Properties;

    /// <summary>
    /// View model for update wizard.
    /// </summary>
    public class UpdateWizardViewModel
    {
        #region private variables

        /// <summary>
        /// Variable for fetch button click command.
        /// </summary>
        private ICommand updateCommand;

        /// <summary>
        /// Variable for cancel click command.
        /// </summary>
        private ICommand cancelCommand;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the FetchClimateViewModel class.
        /// </summary>
        public UpdateWizardViewModel()
        {
            this.updateCommand = new UpdateHandler(this);
            this.cancelCommand = new UpdateCancelHandler(this);

            this.ColumnHeaders = WorkflowController.Instance.GetColumns().ToList();

            this.Input = new UpdateDataViewModel();

            this.Input.AltitudeSytle = AltitudeStyle.Constant;
            this.Input.ColorScheme = Common.ColorScheme.FromData;

            this.Input.DeltaLatitude = 0.0045 / 2;
            this.Input.DeltaLongitude = 0.006 / 2;

            // •	Upper / north boundary Latitude = 38 degrees North = Max Lat 
            // •	Lower / south boundary Latitude = 37.666 degrees North = Min Lat
            // •	Left / west boundary Longitude = -119.6 degrees West = Min Lon
            // •	Right / east boundary Longitude = -119.2 degrees West = Max Lon

            this.Input.MinLatitude = 37.666;
            this.Input.MaxLatitude = 38;
            this.Input.MinLongitude = -119.6;
            this.Input.MaxLongitude = -119.2;
        }

        #endregion

        #region CustomEvent

        /// <summary>
        /// Capture View window close request
        /// </summary>
        public event EventHandler RequestClose;

        #endregion

        #region public variables

        public UpdateDataViewModel Input
        {
            get;
            set;
        }

        public List<string> ColumnHeaders
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the updateCommand.
        /// </summary>
        public ICommand UpdateCommand
        {
            get { return this.updateCommand; }
        }

        /// <summary>
        /// Gets the cancel command.
        /// </summary>
        public ICommand CancelCommand
        {
            get { return this.cancelCommand; }
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

        #region Event handler

        /// <summary>
        /// Cancel button click event handler.
        /// </summary>
        private class UpdateCancelHandler : RelayCommand
        {
            private UpdateWizardViewModel parent;

            /// <summary>
            /// Initializes a new instance of the UpdateCancelHandler class.
            /// </summary>
            /// <param name="parentViewModel">Parent view model.</param>
            public UpdateCancelHandler(UpdateWizardViewModel parentViewModel)
            {
                this.parent = parentViewModel;
            }

            /// <summary>
            /// Execute command event for cancel button.
            /// </summary>
            /// <param name="parameter">Command parameter.</param>
            public override void Execute(object parameter)
            {
                this.parent.OnRequestClose();
            }
        }

        /// <summary>
        /// Fetch button click event handler
        /// </summary>
        private class UpdateHandler : RelayCommand
        {
            private UpdateWizardViewModel parent;

            /// <summary>
            /// Initializes a new instance of the UpdateHandler class.
            /// </summary>
            /// <param name="parentViewModel">Parent view model.</param>
            public UpdateHandler(UpdateWizardViewModel parentViewModel)
            {
                this.parent = parentViewModel;
            }

            /// <summary>
            /// Execute command event for update button.
            /// </summary>
            /// <param name="parameter">Command parameter.</param>
            public override void Execute(object parameter)
            {
                try
                {
                    UpdateDataModel model = new UpdateDataModel();

                    model.AltitudeColumn = this.parent.Input.AltitudeColumn;
                    model.AltitudeSytle = this.parent.Input.AltitudeSytle;
                    model.AltitudeConstant = this.parent.Input.AltitudeConstant;
                    model.AlphaConstant = this.parent.Input.AlphaConstant;
                    model.BetaConstant = this.parent.Input.BetaConstant;

                    model.ColorColumn = this.parent.Input.ColorColumn;
                    model.ColorPalette = this.parent.Input.ColorPalette;
                    model.ColorScheme = this.parent.Input.ColorScheme;
                    model.ColorMax = this.parent.Input.ColorMax;
                    model.ColorMin = this.parent.Input.ColorMin;

                    model.DeltaLatitude = this.parent.Input.DeltaLatitude;
                    model.DeltaLongitude = this.parent.Input.DeltaLongitude;

                    model.RColumn = this.parent.Input.RColumn;
                    model.GColumn = this.parent.Input.GColumn;
                    model.BColumn = this.parent.Input.BColumn;

                    model.MinLatitude = this.parent.Input.MinLatitude;
                    model.MaxLatitude = this.parent.Input.MaxLatitude;
                    model.MinLongitude = this.parent.Input.MinLongitude;
                    model.MaxLongitude = this.parent.Input.MaxLongitude;

                    model.FilterBetweenBoundaries = this.parent.Input.FilterBetweenBoundaries;

                    WorkflowController.Instance.GenerateWWTColumns(model);
                    // WorkflowController.Instance.GenerateHuricaneColumns(model);
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

                // Closing the popup.
                this.parent.OnRequestClose();
            }
        }

        #endregion
    }
}
