//-----------------------------------------------------------------------
// <copyright file="ViewpointViewModel.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Windows.Input;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Class represents the Viewpoint view model.
    /// </summary>
    public class ViewpointViewModel : PropertyChangeBase
    {
        #region Private Properties

        private Perspective perspective;
        private bool isButtonEnabled;
        private bool isSelected;
        private ICommand viewpointNameChangedCommand;
        private ICommand viewpointNameTextChangeCommand;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the ViewpointViewModel class.
        /// </summary>
        /// <param name="perspective">
        /// perspective value
        /// </param>
        public ViewpointViewModel(Perspective perspective)
        {
            this.perspective = perspective;
            this.viewpointNameChangedCommand = new ViewpointNameChangeHandler(this);
            this.viewpointNameTextChangeCommand = new ViewpointNameTextChangeHandler(this);
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
        /// Gets or sets the value of current perspective
        /// </summary>
        public Perspective CurrentPerspective
        {
            get { return perspective; }
            set { perspective = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the Ok button is enabled or not
        /// </summary>
        public bool IsButtonEnabled
        {
            get 
            { 
                return isButtonEnabled; 
            }

            set 
            { 
                isButtonEnabled = value;
                OnPropertyChanged("IsButtonEnabled"); 
            }
        }

        /// <summary>
        /// Gets or sets the value of current perspective name
        /// </summary>
        public string Name
        {
            get
            {
                return this.perspective.Name;
            }

            set
            {
                this.perspective.Name = value;
                OnPropertyChanged("Name"); 
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether values are selected or not 
        /// </summary>
        public bool IsSelected
        {
            get
            {
                return isSelected;
            }

            set
            {
                isSelected = value;
                OnPropertyChanged("IsSelected");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the look at is sky or not
        /// </summary>
        public bool IsSky
        {
            get
            {
                return this.perspective.HasRADec ? true : false;
            }

            set
            {
                this.perspective.HasRADec = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the look  is not sky
        /// </summary>
        public bool IsNotSky
        {
            get
            {
                return this.perspective.HasRADec ? false : true;
            }

            set
            {
                this.perspective.HasRADec = !value;
            }
        }

        /// <summary>
        /// Gets the Viewpoint name change command.
        /// </summary>
        public ICommand ViewpointNameChangeCommand
        {
            get { return this.viewpointNameChangedCommand; }
        }

        /// <summary>
        /// Gets the Viewpoint name text change command.
        /// </summary>
        public ICommand ViewpointNameTextChangeCommand
        {
            get { return this.viewpointNameTextChangeCommand; }
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

        private class ViewpointNameChangeHandler : RelayCommand
        {
            private ViewpointViewModel parent;

            public ViewpointNameChangeHandler(ViewpointViewModel viewpointViewModel)
            {
                this.parent = viewpointViewModel;
            }

            public override void Execute(object parameter)
            {
                if (this.parent != null && parameter != null)
                {
                    string viewpointName = parameter as string;
                    if (!string.IsNullOrWhiteSpace(viewpointName))
                    {
                        this.parent.Name = viewpointName.Trim(); 
                    }

                    // This is OK button click. So request window close
                    this.parent.OnRequestClose();
                }
            }
        }

        private class ViewpointNameTextChangeHandler : RelayCommand
        {
            private ViewpointViewModel parent;

            public ViewpointNameTextChangeHandler(ViewpointViewModel viewpointViewModel)
            {
                this.parent = viewpointViewModel;
            }

            public override void Execute(object parameter)
            {
                if (this.parent != null && parameter != null)
                {
                    string viewpointName = parameter as string;
                    if (string.IsNullOrWhiteSpace(viewpointName))
                    {
                        this.parent.IsButtonEnabled = false;
                    }
                    else
                    {
                        this.parent.IsButtonEnabled = true; 
                    }
                }
            }
        }

        #endregion
    }
}
