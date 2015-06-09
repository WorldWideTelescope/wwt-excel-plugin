//-----------------------------------------------------------------------
// <copyright file="ColumnViewModel.cs" company="Microsoft Corporation">
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
    /// View model for column mapped from WWT and excel header columns
    /// </summary>
    public class ColumnViewModel : PropertyChangeBase
    {
        #region Private Properties

        private ObservableCollection<Column> wwtColumns;
        private Column selectedWWTColumn;
        private string excelHeaderColumn;
        private ICommand mapColumnCommand;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the ColumnViewModel class
        /// </summary>
        public ColumnViewModel()
        {
            this.mapColumnCommand = new MapColumnSelectionChangeHandler(this);
        }

        #endregion

        #region Events

        /// <summary>
        /// Event is fired on map column dropdown change
        /// </summary>
        public event EventHandler MapColumnSelectionChangedEvent;

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the excel column header value
        /// </summary>
        public string ExcelHeaderColumn
        {
            get
            {
                return this.excelHeaderColumn;
            }
            set
            {
                this.excelHeaderColumn = value;
                OnPropertyChanged("ExcelHeaderColumn");
            }
        }

        /// <summary>
        /// Gets or sets the WWT column list
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Setter required for rebind scenarios")]
        public ObservableCollection<Column> WWTColumns
        {
            get
            {
                return this.wwtColumns;
            }
            set
            {
                wwtColumns = value;
                OnPropertyChanged("WWTColumns");
            }
        }

        /// <summary>
        /// Gets or sets the selected WWT column
        /// </summary>
        public Column SelectedWWTColumn
        {
            get
            {
                return this.selectedWWTColumn;
            }
            set
            {
                selectedWWTColumn = value;
                OnPropertyChanged("SelectedWWTColumn");
            }
        }

        #region ICommand

        /// <summary>
        /// Gets map column command for selection changed 
        /// </summary>
        public ICommand MapColumnCommand
        {
            get { return mapColumnCommand; }
        }

        #endregion

        #endregion

        #region Event Handler

        /// <summary>
        /// Event is fired on map column selection changed
        /// </summary>
        private class MapColumnSelectionChangeHandler : RelayCommand
        {
            private ColumnViewModel parent;

            public MapColumnSelectionChangeHandler(ColumnViewModel columnViewModel)
            {
                this.parent = columnViewModel;
            }

            public override void Execute(object parameter)
            {
                if (this.parent != null)
                {
                    this.parent.MapColumnSelectionChangedEvent.OnFire(this.parent, new EventArgs());
                }
            }
        }

        #endregion
    }
}
