//-----------------------------------------------------------------------
// <copyright file="GroupViewModel.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.ObjectModel;
using System.Windows.Input;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// View model for reference frame/layer group population for all layers
    /// </summary>
    public class GroupViewModel : PropertyChangeBase
    {
        #region Private Properties
        private string name;
        private ICommand groupSelectionCommand;

        #endregion

        #region Constructor

        /// <summary>
        ///  Initializes a new instance of the GroupViewModel class 
        /// </summary>
        public GroupViewModel()
        {
            this.ReferenceGroup = new Collection<Group>();
            this.groupSelectionCommand = new GroupSelectionHandler(this);
        }
        #endregion

        #region Events

        /// <summary>
        /// Event is fired on map column dropdown change
        /// </summary>
        public event EventHandler GroupSelectionChangedEvent;

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the name of reference frame
        /// </summary>
        public string Name
        {
            get
            {
                return name;
            }
            set
            {
                name = value;
                OnPropertyChanged("Name");
            }
        }

        /// <summary>
        /// Gets reference groups
        /// </summary>
        public Collection<Group> ReferenceGroup
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the group selection changed command
        /// </summary>
        public ICommand GroupSelectionCommand
        {
            get { return this.groupSelectionCommand; }
        }
        #endregion

        #region Event Handler

        private class GroupSelectionHandler : RelayCommand
        {
            private GroupViewModel parent;
            public GroupSelectionHandler(GroupViewModel groupViewModel)
            {
                this.parent = groupViewModel;
            }

            public override void Execute(object parameter)
            {
                if (this.parent != null && parameter != null)
                {
                    this.parent.GroupSelectionChangedEvent.OnFire(parameter, new EventArgs());
                }
            }
        }
        #endregion
    }
}
