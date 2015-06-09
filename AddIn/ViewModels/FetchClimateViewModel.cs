// -----------------------------------------------------------------------
// <copyright file="FetchClimateViewModel.cs" company="AditiTechnologies Pvt Ltd">
//  View model for fetch climate view.
// </copyright>
// -----------------------------------------------------------------------

namespace Microsoft.Research.Wwt.Excel.Addin
{
    using System;
    using System.Collections.Generic;
    using System.Windows.Input;
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Research.Wwt.Excel.Common;

    /// <summary>
    /// View model for fetch climate view.
    /// </summary>
    public class FetchClimateViewModel
    {
        #region private variables

        /// <summary>
        /// Variable for fetch button click command.
        /// </summary>
        private ICommand fetchClimateCommand;

        /// <summary>
        /// Variable for cancel click command.
        /// </summary>
        private ICommand cancelCommand;

        /// <summary>
        /// Variable for fetch climate input model class.
        /// </summary>
        private FetchClimateInputModel fetchClimateInputs;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the FetchClimateViewModel class.
        /// </summary>
        public FetchClimateViewModel()
        {
            this.fetchClimateCommand = new FetchClimateHandler(this);
            this.cancelCommand = new CancelEventHandler(this);
            this.fetchClimateInputs = new FetchClimateInputModel();

            // Assigning default values for testing
            this.fetchClimateInputs.MinLatitude = 39.80710;
            this.fetchClimateInputs.MaxLatitude = 48.34040;
            this.fetchClimateInputs.MinLongitude = -129.35869;
            this.fetchClimateInputs.MaxLongitude = -114.68096;
            this.fetchClimateInputs.DeltaLatitude = 0.3;
            this.fetchClimateInputs.DeltaLongitude = 0.3;
        }

        #endregion     

        #region CustomEvent

        /// <summary>
        /// Capture View window close request
        /// </summary>
        public event EventHandler RequestClose;

        #endregion

        #region public variables

        /// <summary>
        /// Gets the fetch climate command.
        /// </summary>
        public ICommand FetchClimateCommand
        {
            get { return this.fetchClimateCommand; }
        }

        /// <summary>
        /// Gets the cancel command.
        /// </summary>
        public ICommand CancelCommand
        {
            get { return this.cancelCommand; }
        }

        /// <summary>
        ///  Gets or sets fetch climate input values.
        /// </summary>
        public FetchClimateInputModel FetchClimateInputs
        {
            get
            {
                return this.fetchClimateInputs;
            }
            set
            {
                this.fetchClimateInputs = value;
            }
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
        private class CancelEventHandler : RelayCommand
        {
            private FetchClimateViewModel parent;

            /// <summary>
            /// Initializes a new instance of the CancelEventHandler class.
            /// </summary>
            /// <param name="fetchClimateViewModel">Parent view model.</param>
            public CancelEventHandler(FetchClimateViewModel fetchClimateViewModel)
            {
                this.parent = fetchClimateViewModel;
            }

            /// <summary>
            /// Execute command event for fetch climate button.
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
        private class FetchClimateHandler : RelayCommand
        {
            private FetchClimateViewModel parent;

            /// <summary>
            /// Initializes a new instance of the FetchClimateHandler class.
            /// </summary>
            /// <param name="fetchClimateViewModel">Parent view model.</param>
            public FetchClimateHandler(FetchClimateViewModel fetchClimateViewModel)
            {
                this.parent = fetchClimateViewModel;
            }

            /// <summary>
            /// Execute command event for fetch climate button.
            /// </summary>
            /// <param name="parameter">Command parameter.</param>
            public override void Execute(object parameter)
            {
                WorkflowController.Instance.InsertFetchClimateData(this.parent.FetchClimateInputs.MinLatitude, this.parent.FetchClimateInputs.MaxLatitude, this.parent.FetchClimateInputs.MinLongitude, this.parent.FetchClimateInputs.MaxLongitude, this.parent.FetchClimateInputs.DeltaLatitude, this.parent.FetchClimateInputs.DeltaLongitude);

                // Closing the popup.
                this.parent.OnRequestClose();
            }
        }

        #endregion
    }
}
