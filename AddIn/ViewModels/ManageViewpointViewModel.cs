//-----------------------------------------------------------------------
// <copyright file="ManageViewpointViewModel.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Windows.Input;
using Microsoft.Research.Wwt.Excel.Common; 

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Class represents the Manage Viewpoint view model.
    /// </summary>
    public class ManageViewpointViewModel : PropertyChangeBase
    {
        #region Private Properties

        private ObservableCollection<ViewpointViewModel> allViewpoints;
        private bool isSelected;
        private ICommand selectionChangedCommand;
        private ICommand renameViewpointCommand;
        private ICommand deleteViewpointCommand;
        private ICommand gotoViewpointCommand;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the ManageViewpointViewModel class.
        /// </summary>
        /// <param name="allViewpoints">
        /// all allViewpoints 
        /// </param>
        public ManageViewpointViewModel(ObservableCollection<Perspective> allViewpoints)
        {
            this.allViewpoints = new ObservableCollection<ViewpointViewModel>();

            if (allViewpoints != null)
            {
                foreach (Perspective perspective in allViewpoints)
                {
                    this.allViewpoints.Add(new ViewpointViewModel(perspective));
                }
            }

            this.selectionChangedCommand = new SelectionChangeHandler(this);
            this.renameViewpointCommand = new RenameViewpointHandler(this);
            this.deleteViewpointCommand = new DeleteViewpointHandler(this);
            this.gotoViewpointCommand = new GotoViewpointHandler(this);
        }

        #endregion

        #region CustomEvent

        /// <summary>
        /// Delete Viewpoint Event
        /// </summary>
        public event EventHandler DeleteViewpointEvent;
        
        /// <summary>
        /// Rename Viewpoint Event
        /// </summary>
        public event EventHandler RenameViewpointEvent;
        
        /// <summary>
        /// Go to Viewpoint Event
        /// </summary>
        public event EventHandler GotoViewpointEvent;

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the value of all Viewpoints
        /// </summary>
        [SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Observable collection for binding in XAML")]
        public ObservableCollection<ViewpointViewModel> AllViewpoint
        {
            get { return allViewpoints; }
            set { allViewpoints = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether a record is selected or not
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
        /// Gets the selection change command.
        /// </summary>
        public ICommand SelectionChangedCommand
        {
            get { return this.selectionChangedCommand; }
        }

        /// <summary>
        /// Gets the Rename Viewpoint command.
        /// </summary>
        public ICommand RenameViewpointCommand
        {
            get { return this.renameViewpointCommand; }
        }

        /// <summary>
        /// Gets the Delete Viewpoint command.
        /// </summary>
        public ICommand DeleteViewpointCommand
        {
            get { return this.deleteViewpointCommand; }
        }

        /// <summary>
        /// Gets the go to Viewpoint command.
        /// </summary>
        public ICommand GotoViewpointCommand
        {
            get { return this.gotoViewpointCommand; }
        }

        #endregion

        #region Event Handler

        private class SelectionChangeHandler : RelayCommand
        {
            private ManageViewpointViewModel parent;

            public SelectionChangeHandler(ManageViewpointViewModel manageViewpointViewModel)
            {
                this.parent = manageViewpointViewModel;
            }

            public override void Execute(object parameter)
            {
                if (this.parent != null && parameter != null)
                {
                    if (this.parent.AllViewpoint.Any() && this.parent.AllViewpoint.Where(item => item.IsSelected == true).FirstOrDefault() != null)
                    {
                        this.parent.IsSelected = true;
                    }
                    else
                    {
                        this.parent.IsSelected = false;
                    }
                }
                else
                {
                    this.parent.IsSelected = false;
                }
            }
        }

        private class RenameViewpointHandler : RelayCommand
        {
            private ManageViewpointViewModel parent;

            public RenameViewpointHandler(ManageViewpointViewModel manageViewpointViewModel)
            {
                this.parent = manageViewpointViewModel;
            }

            public override void Execute(object parameter)
            {
                if (this.parent != null && parameter != null)
                {
                    this.parent.RenameViewpointEvent.OnFire(parameter, new EventArgs());
                }
            }
        }

        private class DeleteViewpointHandler : RelayCommand
        {
            private ManageViewpointViewModel parent;

            public DeleteViewpointHandler(ManageViewpointViewModel manageViewpointViewModel)
            {
                this.parent = manageViewpointViewModel;
            }

            public override void Execute(object parameter)
            {
                if (this.parent != null && parameter != null)
                {
                    this.parent.DeleteViewpointEvent.OnFire(parameter, new EventArgs());
                }
            }
        }

        private class GotoViewpointHandler : RelayCommand
        {
            private ManageViewpointViewModel parent;

            public GotoViewpointHandler(ManageViewpointViewModel manageViewpointViewModel)
            {
                this.parent = manageViewpointViewModel;
            }

            public override void Execute(object parameter)
            {
                if (this.parent != null && parameter != null)
                {
                    this.parent.GotoViewpointEvent.OnFire(parameter, new EventArgs());
                }
            }
        }

        #endregion
    }
}
