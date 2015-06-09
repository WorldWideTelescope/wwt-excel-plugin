//-----------------------------------------------------------------------
// <copyright file="LayerMapDropDownViewModel.cs" company="Microsoft Corporation">
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
    /// View model for layer drop down
    /// </summary>
    public class LayerMapDropDownViewModel : PropertyChangeBase
    {
        #region Private Properties
        private string name;
        private Collection<GroupChildren> groupCollection;
        private ICommand selectionCommand;
        private string id;
        #endregion

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the LayerMapDropDownViewModel class.
        /// </summary>
        public LayerMapDropDownViewModel()
        {
            this.GroupCollection = new Collection<GroupChildren>();
            this.selectionCommand = new LayerSelectionHandler(this);
        }
        #endregion

        #region Events

        /// <summary>
        /// Event is fired on map column dropdown change
        /// </summary>
        public event EventHandler LayerSelectionChangedEvent;

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets Id for the group
        /// </summary>
        public string ID
        {
            get
            {
                return this.id;
            }
            set
            {
                this.id = value;
                OnPropertyChanged("ID");
            }
        }

        /// <summary>
        /// Gets or sets the name of group
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
                OnPropertyChanged("Name");
            }
        }

        /// <summary>
        /// Gets or sets group collection
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Observable collection for binding in XAML")]
        public Collection<GroupChildren> GroupCollection
        {
            get
            {
                return groupCollection;
            }
            set
            {
                groupCollection = value;
                OnPropertyChanged("GroupCollection");
            }
        }

        #region ICommand

        public ICommand LayerSelectionCommand
        {
            get
            {
                return this.selectionCommand;
            }
        }

        #endregion
        #endregion

        #region Event Handler

        /// <summary>
        /// This is the implementation of the relay command class for layer selection event.
        /// </summary>
        private class LayerSelectionHandler : RelayCommand
        {
            private LayerMapDropDownViewModel parent;

            /// <summary>
            /// Initializes a new instance of the LayerSelectionHandler class.
            /// </summary>
            /// <param name="layerMapDropDownViewModel">
            /// layerMapDropDownViewModel instance.
            /// </param>
            public LayerSelectionHandler(LayerMapDropDownViewModel layerMapDropDownViewModel)
            {
                this.parent = layerMapDropDownViewModel;
            }

            /// <summary>
            /// This method is called when the control to which a command is bound
            /// is actuated by the user
            /// </summary>
            /// <param name="parameter">
            /// Command parameter
            /// </param>
            public override void Execute(object parameter)
            {
                if (this.parent != null && parameter != null)
                {
                    this.parent.LayerSelectionChangedEvent.OnFire(parameter, new EventArgs());
                }
            }
        }

        #endregion
    }
}