//-----------------------------------------------------------------------
// <copyright file="Column.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Collections.ObjectModel;

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Class that will represent the column that need to be mapped
    /// </summary>
    public class Column
    {
        private ColumnType columnType;
        private string columnDisplayValue;
        private Collection<string> columnMatchValues;

        /// <summary>
        /// Initializes a new instance of the Column class.
        /// </summary>
        public Column(ColumnType columnType, string columnDisplayValue, Collection<string> columnComparisonValue)
        {
            this.columnType = columnType;
            this.columnDisplayValue = columnDisplayValue;
            this.columnMatchValues = columnComparisonValue;
        }

        /// <summary>
        /// Gets or sets Column Type 
        /// </summary>
        public ColumnType ColType
        {
            get { return columnType; }
            set { columnType = value; }
        }

        /// <summary>
        /// Gets or sets Column Display Value
        /// </summary>
        public string ColumnDisplayValue
        {
            get { return columnDisplayValue; }
            set { columnDisplayValue = value; }
        }

        /// <summary>
        /// Gets Column Match Value collection
        /// </summary>
        public ReadOnlyCollection<string> ColumnMatchValues
        {
            get { return new ReadOnlyCollection<string>(columnMatchValues); }
        }
    }
}
