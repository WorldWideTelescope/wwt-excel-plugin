//-----------------------------------------------------------------------
// <copyright file="LayerMap.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;
using System.Threading;
using System.Windows.Threading;
using Microsoft.Office.Interop.Excel;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// All properties required to persist a layer and use it as a model for Layer manager
    /// As this class is de-serialized, need to ensure that all the properties are available after de-serialization as well
    /// Also use public properties and not private properties in this class as de-serialization will not set private properties
    /// Build public properties dynamically if they are not serialized and saved in the workbook
    /// </summary>
    [DataContract]
    internal class LayerMap
    {
        #region Private members

        /// <summary>
        /// Collection of all columns
        /// </summary>
        private Collection<Column> columnsList;

        /// <summary>
        ///  Collection of header row data
        /// </summary>
        private Collection<string> headerRowData;

        /// <summary>
        /// Range name for this layer
        /// </summary>
        private Name rangeName;

        /// <summary>
        /// Name for layer
        /// </summary>
        private string name;

        /// <summary>
        /// Indicates whether the local Layer is in sync with WWT layer or not.
        /// This is required in workbook load scenario where the layer id is same as the id in WWT but
        /// the layer properties might be different because of disconnected mode.
        /// </summary>
        private bool isNotInSync;

        /// <summary>
        /// Token for cancelling the notification request.
        /// </summary>
        private CancellationTokenSource cancellationTokenSource;

        #endregion Private members

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the LayerMap class.
        /// </summary>
        /// <param name="name">
        /// Range for which the layer.
        /// </param>
        internal LayerMap(Name name)
        {
            this.rangeName = name;
            this.RangeDisplayName = name.Name;
            this.RangeAddress = name.RefersTo as string;
            this.MapType = LayerMapType.Local;
            this.columnsList = ColumnExtensions.PopulateColumnList();
            this.SetAutoMap();
            this.LayerDetails = new Layer();
            this.LayerDetails.Name = this.RangeDisplayName;
            this.SetLayerProperties();
        }

        /// <summary>
        /// Initializes a new instance of the LayerMap class.
        /// </summary>
        /// <param name="layer">
        /// Layer details.
        /// </param>
        internal LayerMap(Layer layer)
        {
            this.LayerDetails = layer;
            this.MapType = LayerMapType.WWT;
            this.columnsList = ColumnExtensions.PopulateColumnList();

            // Set mapped column type based on header row data
            SetMappedColumnType();
            
            // For WWT layers, notification will be started immediately.
            StartNotifying();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets layer name
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
            }
        }

        /// <summary>
        /// Gets the Name property for the range.
        /// </summary>
        internal Name RangeName
        {
            get
            {
                if (this.rangeName == null)
                {
                    if(WorkbookReference == null)
                        this.rangeName = ThisAddIn.ExcelApplication.ActiveWorkbook.Names.GetNamedRange(this.RangeDisplayName);
                    else
                        this.rangeName = WorkbookReference.Names.GetNamedRange(this.RangeDisplayName);
                }

                return this.rangeName;
            }
        }

        /// <summary>
        /// Temp reference. See - this.RangeName & WorkbookMap.LocalLayerMaps
        /// </summary>
        internal Workbook WorkbookReference { get; set; }

        /// <summary>
        /// Gets or sets the column collection.
        /// </summary>
        internal Collection<Column> ColumnsList
        {
            get
            {
                this.columnsList = this.columnsList ?? ColumnExtensions.PopulateColumnList();
                return this.columnsList;
            }
            set
            {
                this.columnsList = value;
            }
        }

        /// <summary>
        /// Gets the value of RangeName.
        /// Set on create
        /// </summary>
        [DataMember]
        internal string RangeDisplayName { get; private set; }

        /// <summary>
        /// Gets or sets the value of range address.
        /// </summary>
        [DataMember]
        internal string RangeAddress { get; set; }

        /// <summary>
        /// Gets or sets the value of Layer
        /// Layer mappings, properties and marker properties
        /// </summary>
        [DataMember]
        internal Layer LayerDetails { get; set; }

        /// <summary>
        /// Gets a value indicating whether the Layer is valid or not.
        /// Need not be serialized, just loaded and set
        /// </summary>
        internal bool IsValid
        {
            get
            {
                return this.RangeName.IsValid();
            }
        }

        /// <summary>
        /// Gets or sets a value indicating the Layer type.
        /// </summary>
        [DataMember]
        internal LayerMapType MapType { get; set; }

        /// <summary>
        /// Gets or sets the Data in first row of the selection range
        /// </summary>
        internal Collection<string> HeaderRowData
        {
            get
            {
                if (this.headerRowData == null)
                {
                    if (this.RangeName.IsValid())
                    {
                        this.headerRowData = this.RangeName.RefersToRange.GetHeader();
                    }
                }

                return this.headerRowData;
            }
            set
            {
                this.headerRowData = value;
            }
        }

        /// <summary>
        /// Gets or sets mapped data for first row data in the selection range
        /// Need to be serialized and saved
        /// </summary>
        [DataMember]
        internal Collection<ColumnType> MappedColumnType { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the local Layer is in sync with WWT layer or not.
        /// This is required in workbook load scenario where the layer id is same as the id in WWT but
        /// the layer properties might be different because of disconnected mode.
        /// </summary>
        internal bool IsNotInSync
        {
            get
            {
                return isNotInSync;
            }
            set
            {
                isNotInSync = value;

                if (!isNotInSync)
                {
                    // For LocalInWWT and WWT layers only, need to start the notifications based on IsNotInSync value.
                    if (this.MapType == LayerMapType.LocalInWWT || this.MapType == LayerMapType.WWT)
                    {
                        StartNotifying();
                    }
                }
                else
                {
                    // Even for Local layers, need to stop the notifications based on IsNotInSync value, since there could be a scenario in which
                    // notifications started when the layer was LocalInWWT or WWT, but later changed as Local. This will happen only while
                    // running automated unit test cases.
                    StopNotifying();
                }
            }
        }

        /// <summary>
        /// Gets or sets the date time when visualize was clicked
        
        /// </summary>
        internal DateTime VisualizeClickTime { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether Visualize button was clicked
      
        /// </summary>
        internal bool IsVisualizeClicked { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the layer properties are changed through code or not.
        /// </summary>
        internal bool IsPropertyChangedFromCode { get; set; }

        #endregion

        #region Internal methods

        /// <summary>
        /// Set MappedColumnData based on HeaderRowData
        /// </summary>
        internal void SetAutoMap()
        {
            MappedColumnType = new Collection<ColumnType>();

            // Boolean variables to check whether depth columns are having exact match or contains match. Exact match will have
            // precedence over contains match.
            bool depthExactMatch = false, depthContainsMatch = false;

            // Boolean variables to check whether Lat/long columns are having exact match or contains match. Exact match will have
            // precedence over contains match.
            bool latLonExactMatch = false, latLonContainsMatch = false;

            bool startDateContainsMatch = false, isXColSelected = false, isYColSelected = false, isZColSelected = false, radecMatch = false;
            int mappedColumnsCount = 0;

            // Initialize MappedColumnType with None column type.
            for (int i = 0; i < this.HeaderRowData.Count; i++)
            {
                MappedColumnType.Add(ColumnType.None);
            }

            // Start from second column since first column is select one.
            for (int i = 1; i < this.ColumnsList.Count; i++)
            {
                if (mappedColumnsCount == this.HeaderRowData.Count)
                {
                    // If all the headers are mapped, no need to proceed with other columns.
                    break;
                }

                bool mappingFound = false;
                Column column = this.columnsList[i];

                // 1. If the current column is depth column (Depth/Altitude/Distance) and depth is already having exact match mapping, ignore this column.
                // 2. If the current column is XYZ column and one of X or Y or Z is already mapped, ignore this column
                // 3. If the column is RA/Dec, but Lat/Lon columns are already mapped with Exact match, ignore this column.
                // 4. If the column is RA/Dec and SetAutoMap is called for an existing layer which belong to Earth reference frame, ignore this column.
                // 5. If the column is Lat/Long and SetAutoMap is called for an existing layer which belong to Sky reference frame, ignore this column.
                if ((column.IsDepthColumn() && depthExactMatch) ||
                        ((column.ColType == ColumnType.X || column.ColType == ColumnType.ReverseX) && (isXColSelected || latLonExactMatch || radecMatch)) ||
                        ((column.ColType == ColumnType.Y || column.ColType == ColumnType.ReverseY) && (isYColSelected || latLonExactMatch || radecMatch)) ||
                        ((column.ColType == ColumnType.Z || column.ColType == ColumnType.ReverseZ) && (isZColSelected || latLonExactMatch || radecMatch)) ||
                        ((column.ColType == ColumnType.RA || column.ColType == ColumnType.Dec) && (latLonExactMatch || (this.LayerDetails != null && this.LayerDetails.Group.IsPlanet()))) ||
                        ((column.ColType == ColumnType.Lat || column.ColType == ColumnType.Long) && (this.LayerDetails != null && !this.LayerDetails.Group.IsPlanet())))
                {
                    continue;
                }

                // First look for exact match
                foreach (string columnMatchValue in column.ColumnMatchValues)
                {
                    for (int j = 0; j < this.HeaderRowData.Count; j++)
                    {
                        // If the mapped column type is start date and if the start date is contains match, then exact match check on end date 
                        // has to be done as start date and end date have common column match values.
                        if (MappedColumnType[j] != ColumnType.None && !(MappedColumnType[j] == ColumnType.StartDate && startDateContainsMatch))
                        {
                            // This header is already mapped with some other column, continue with other headers.
                            continue;
                        }

                        string currentHeader = this.HeaderRowData[j].ToUpper(CultureInfo.CurrentCulture).Trim();

                        if (currentHeader.Equals(columnMatchValue.ToUpper(CultureInfo.CurrentCulture)))
                        {
                            if (column.IsDepthColumn())
                            {
                                // Remove the existing contains matches if any exists for depth columns (Depth/Altitude/Distance).
                                if (depthContainsMatch)
                                {
                                    RemoveDepthColumnMappings();
                                }

                                depthExactMatch = true;
                            }

                            MappedColumnType[j] = column.ColType;
                            mappingFound = true;
                            mappedColumnsCount++;

                            switch (column.ColType)
                            {
                                case ColumnType.Lat:
                                case ColumnType.Long:
                                    latLonExactMatch = true;
                                    break;
                                case ColumnType.RA:
                                case ColumnType.Dec:
                                    radecMatch = true;

                                    // If RA/Dec got exact match, remove the mappings of Lat/Long which are found with Contains match, if any.
                                    // If Lat/Long has exact match, this code will not be hit since RA/Dec are ignored in case of Lat/Lon exact match above.
                                    RemoveLatLongMappings();
                                    break;
                                case ColumnType.X:
                                case ColumnType.ReverseX:
                                    RemoveLatLongMappings();
                                    RemoveRaDecMappings();
                                    isXColSelected = true;
                                    break;
                                case ColumnType.Y:
                                case ColumnType.ReverseY:
                                    RemoveLatLongMappings();
                                    RemoveRaDecMappings();
                                    isYColSelected = true;
                                    break;
                                case ColumnType.Z:
                                case ColumnType.ReverseZ:
                                    RemoveLatLongMappings();
                                    RemoveRaDecMappings();
                                    isZColSelected = true;
                                    break;
                                default:
                                    break;
                            }

                            break;
                        }
                    }

                    if (mappingFound)
                    {
                        // Got a mapping for the current Column, break the loop of Column headers.
                        break;
                    }
                }

                if (!mappingFound && !column.IsXYZColumn())
                {
                    if ((column.ColType == ColumnType.Dec && latLonContainsMatch) || (column.IsDepthColumn() && depthContainsMatch))
                    {
                        // 1. If Lat/Long columns are already mapped with contains match, Dec should not be mapped again.
                        //    RA needs to be mapped, since Dec can have Exact match, since Dec mapping is done last.
                        // 2. If Depth (Depth/Altitude/Distance) columns are already mapped with contains match, ignore the current Depth column.
                        continue;
                    }

                    // Look for contains match, if no mapping found in Exact match
                    foreach (string columnMatchValue in column.ColumnMatchValues)
                    {
                        // If column type is RA or DEC, contains match should not be done for match values "RA" and "Dec", since they
                        // are small words can be found in many words.
                        if ((column.ColType == ColumnType.RA && columnMatchValue.Equals(Properties.Resources.RADisplayValue)) ||
                                (column.ColType == ColumnType.Dec && columnMatchValue.Equals(Properties.Resources.DecDisplayValue)))
                        {
                            continue;
                        }

                        for (int j = 0; j < this.HeaderRowData.Count; j++)
                        {
                            if (MappedColumnType[j] != ColumnType.None)
                            {
                                // This header is already mapped with some other column, continue with other headers.
                                continue;
                            }

                            string currentHeader = this.HeaderRowData[j].ToUpper(CultureInfo.CurrentCulture).Trim();

                            if (currentHeader.Contains(columnMatchValue.ToUpper(CultureInfo.CurrentCulture)))
                            {
                                MappedColumnType[j] = column.ColType;
                                mappingFound = true;
                                mappedColumnsCount++;
                                if (column.IsDepthColumn())
                                {
                                    depthContainsMatch = true;
                                }

                                if (column.ColType == ColumnType.Lat || column.ColType == ColumnType.Long)
                                {
                                    latLonContainsMatch = true;
                                }
                                else if (column.ColType == ColumnType.RA || column.ColType == ColumnType.Dec)
                                {
                                    radecMatch = true;
                                }
                                else if (column.ColType == ColumnType.StartDate)
                                {
                                    // If the start date is a contains check then end date need to be checked as well for a exact match 
                                    // as both have some common match values.
                                    startDateContainsMatch = true;
                                    mappedColumnsCount--;
                                }

                                break;
                            }
                        }

                        if (mappingFound)
                        {
                            // Got a mapping for the current Column, break the loop of Column headers.
                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Set MappedColumnData based on Layer details
        /// </summary>
        internal void SetMappedColumnType()
        {
            if (this.LayerDetails != null && !string.IsNullOrEmpty(this.LayerDetails.ID))
            {
                this.HeaderRowData = WWTManager.GetLayerHeader(this.LayerDetails.ID);
                MappedColumnType = new Collection<ColumnType>();

                // Get all mapped columns first
                for (int count = 0; count < this.HeaderRowData.Count; count++)
                {
                    MappedColumnType.Add(ColumnType.None);
                }

                if (this.LayerDetails.LatColumn > Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.LatColumn)
                {
                    MappedColumnType[this.LayerDetails.LatColumn] = ColumnType.Lat;
                }
                if (this.LayerDetails.LngColumn > Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.LngColumn)
                {
                    MappedColumnType[this.LayerDetails.LngColumn] = ColumnType.Long;
                }
                if (this.LayerDetails.GeometryColumn > Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.GeometryColumn)
                {
                    MappedColumnType[this.LayerDetails.GeometryColumn] = ColumnType.Geo;
                }
                if (this.LayerDetails.ColorMapColumn > Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.ColorMapColumn)
                {
                    MappedColumnType[this.LayerDetails.ColorMapColumn] = ColumnType.Color;
                }
                if (this.LayerDetails.AltColumn > Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.AltColumn)
                {
                    switch (this.LayerDetails.AltType)
                    {
                        case AltType.Altitude:
                            MappedColumnType[this.LayerDetails.AltColumn] = ColumnType.Alt;
                            break;
                        case AltType.Depth:
                            MappedColumnType[this.LayerDetails.AltColumn] = ColumnType.Depth;
                            break;
                        case AltType.Distance:
                            MappedColumnType[this.LayerDetails.AltColumn] = ColumnType.Distance;
                            break;
                    }
                }

                if (this.LayerDetails.StartDateColumn > Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.StartDateColumn)
                {
                    MappedColumnType[this.LayerDetails.StartDateColumn] = ColumnType.StartDate;
                }
                if (this.LayerDetails.EndDateColumn > Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.EndDateColumn)
                {
                    MappedColumnType[this.LayerDetails.EndDateColumn] = ColumnType.EndDate;
                }
                if (this.LayerDetails.RAColumn > Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.RAColumn)
                {
                    MappedColumnType[this.LayerDetails.RAColumn] = ColumnType.RA;
                }
                if (this.LayerDetails.DecColumn > Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.DecColumn)
                {
                    MappedColumnType[this.LayerDetails.DecColumn] = ColumnType.Dec;
                }
                if (this.LayerDetails.XAxis > Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.XAxis)
                {
                    MappedColumnType[this.LayerDetails.XAxis] = ColumnType.X;
                }
                if (this.LayerDetails.YAxis > Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.YAxis)
                {
                    MappedColumnType[this.LayerDetails.YAxis] = ColumnType.Y;
                }
                if (this.LayerDetails.ZAxis > Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.ZAxis)
                {
                    MappedColumnType[this.LayerDetails.ZAxis] = ColumnType.Z;
                }
                if (this.LayerDetails.ReverseXAxis && this.LayerDetails.XAxis > Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.XAxis)
                {
                    MappedColumnType[this.LayerDetails.XAxis] = ColumnType.ReverseX;
                }
                if (this.LayerDetails.ReverseYAxis && this.LayerDetails.YAxis > Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.YAxis)
                {
                    MappedColumnType[this.LayerDetails.YAxis] = ColumnType.ReverseY;
                }
                if (this.LayerDetails.ReverseZAxis && this.LayerDetails.ZAxis > Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.ZAxis)
                {
                    MappedColumnType[this.LayerDetails.ZAxis] = ColumnType.ReverseZ;
                }

                // Sets size column mapping
                SetMappingOnSizeColumn();
            }
        }

        /// <summary>
        /// Set Layer properties based on auto mappings done
        /// </summary>
        internal void SetLayerProperties()
        {
            this.SetLayerColumnProperties();

            // If magnitude column is mapped, use index of magnitude column, else set to none
            var magColumn = MappedColumnType.Where(item => item == ColumnType.Mag).FirstOrDefault();
            if (magColumn != ColumnType.None)
            {
                this.LayerDetails.SizeColumn = MappedColumnType.IndexOf(magColumn);
                this.LayerDetails.ScaleFactor = 1;
            }
            else
            {
                this.LayerDetails.SizeColumn = Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex;
                this.LayerDetails.ScaleFactor = Microsoft.Research.Wwt.Excel.Common.Constants.DefaultScaleFactor;
            }

            // Sets the name column to first value in the sheet
            this.LayerDetails.NameColumn = Microsoft.Research.Wwt.Excel.Common.Constants.DefaultNameColumnIndex;

            // Set the scale type based on the group the layer belongs to.
            this.LayerDetails.PointScaleType = this.LayerDetails.Group.IsPlanet() ? ScaleType.Power : ScaleType.StellarMagnitude;
        }

        /// <summary>
        /// Sets layer column properties based on mapping
        /// </summary>
        internal void SetLayerColumnProperties()
        {
            // Reset all the column properties
            this.LayerDetails.InitializeColumnDefaults();

            // Set it based on latest column mappings
            foreach (ColumnType colType in this.MappedColumnType)
            {
                if (colType != ColumnType.None)
                {
                    switch (colType)
                    {
                        case ColumnType.Lat:
                            this.LayerDetails.LatColumn = this.MappedColumnType.IndexOf(colType);
                            break;

                        case ColumnType.Long:
                            this.LayerDetails.LngColumn = this.MappedColumnType.IndexOf(colType);
                            break;

                        case ColumnType.Alt:
                            this.LayerDetails.AltColumn = this.MappedColumnType.IndexOf(colType);
                            this.LayerDetails.AltType = AltType.Altitude;
                            this.LayerDetails.AltUnit = AltUnit.Meters;
                            break;

                        case ColumnType.Depth:
                            this.LayerDetails.AltColumn = this.MappedColumnType.IndexOf(colType);
                            this.LayerDetails.AltType = AltType.Depth;
                            this.LayerDetails.AltUnit = AltUnit.Kilometers;
                            break;

                        case ColumnType.Distance:
                            this.LayerDetails.AltColumn = this.MappedColumnType.IndexOf(colType);
                            this.LayerDetails.AltType = AltType.Distance;
                            this.LayerDetails.AltUnit = AltUnit.Meters;
                            break;

                        case ColumnType.Color:
                            this.LayerDetails.ColorMapColumn = this.MappedColumnType.IndexOf(colType);
                            break;

                        case ColumnType.EndDate:
                            this.LayerDetails.EndDateColumn = this.MappedColumnType.IndexOf(colType);
                            break;

                        case ColumnType.StartDate:
                            this.LayerDetails.StartDateColumn = this.MappedColumnType.IndexOf(colType);
                            break;

                        case ColumnType.Geo:
                            this.LayerDetails.GeometryColumn = this.MappedColumnType.IndexOf(colType);
                            break;
                        case ColumnType.RA:
                            this.LayerDetails.RAColumn = this.MappedColumnType.IndexOf(colType);
                            break;
                        case ColumnType.Dec:
                            this.LayerDetails.DecColumn = this.MappedColumnType.IndexOf(colType);
                            break;
                        case ColumnType.X:
                            this.LayerDetails.XAxis = this.MappedColumnType.IndexOf(colType);
                            break;
                        case ColumnType.ReverseX:
                            this.LayerDetails.XAxis = this.MappedColumnType.IndexOf(colType);
                            this.LayerDetails.ReverseXAxis = true;
                            break;
                        case ColumnType.Y:
                            this.LayerDetails.YAxis = this.MappedColumnType.IndexOf(colType);
                            break;
                        case ColumnType.ReverseY:
                            this.LayerDetails.YAxis = this.MappedColumnType.IndexOf(colType);
                            this.LayerDetails.ReverseYAxis = true;
                            break;
                        case ColumnType.Z:
                            this.LayerDetails.ZAxis = this.MappedColumnType.IndexOf(colType);
                            break;
                        case ColumnType.ReverseZ:
                            this.LayerDetails.ZAxis = this.MappedColumnType.IndexOf(colType);
                            this.LayerDetails.ReverseZAxis = true;
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// Resets the range properties for the layer
        /// </summary>
        /// <param name="resetRangeName">Range names</param>
        internal void ResetRange(Name resetRangeName)
        {
            this.rangeName = resetRangeName;
            this.RangeDisplayName = resetRangeName.Name;
            this.RangeAddress = resetRangeName.RefersTo as string;
        }

        #endregion

        #region Private method

        /// <summary>
        /// Sets mapping on size column 
        /// If size column is set to a column which doesn't have any mapping (Select One), the mapping for
        /// column is set to "Magnitude"
        /// </summary>
        private void SetMappingOnSizeColumn()
        {
            if (this.LayerDetails.SizeColumn != Microsoft.Research.Wwt.Excel.Common.Constants.DefaultColumnIndex && MappedColumnType.Count > this.LayerDetails.SizeColumn
                && MappedColumnType[this.LayerDetails.SizeColumn] == ColumnType.None)
            {
                MappedColumnType[this.LayerDetails.SizeColumn] = ColumnType.Mag;
            }
        }

        /// <summary>
        /// Starts notifying if any change to this layer's properties in WWT. WWT will send notification which needs to be
        /// handled by the LayerMap object.
        /// </summary>
        private void StartNotifying()
        {
            if (this.cancellationTokenSource == null)
            {
                using (BackgroundWorker propertyChangeNotifier = new BackgroundWorker())
                {
                    // Initialize the background worker thread which will be listening to WWT for notifications.
                    propertyChangeNotifier.WorkerSupportsCancellation = true;
                    propertyChangeNotifier.DoWork += new DoWorkEventHandler(OnPropertyChangeNotifierDoWork);
                    propertyChangeNotifier.RunWorkerCompleted += new RunWorkerCompletedEventHandler(OnPropertyChangeNotifierCompleted);
                    this.cancellationTokenSource = new CancellationTokenSource();
                    IsPropertyChangedFromCode = false;
                    propertyChangeNotifier.RunWorkerAsync();
                }
            }
        }

        /// <summary>
        /// Stop notifying the layer property changes.
        /// </summary>
        private void StopNotifying()
        {
            if (this.cancellationTokenSource != null)
            {
                this.cancellationTokenSource.Cancel();
                this.cancellationTokenSource.Dispose();
                this.cancellationTokenSource = null;
            }
        }

        /// <summary>
        /// Implementation for the BackgroundWorker which will be syncing the layer properties.
        /// </summary>
        /// <param name="sender">Background Worker</param>
        /// <param name="e">BackgroundWorker event arguments</param>
        private void OnPropertyChangeNotifierDoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                e.Result = false;

                int layerVersion = WWTManager.CreateLayerNotification(this.LayerDetails.ID, this.LayerDetails.Version, this.cancellationTokenSource.Token);
                WorkbookMap parentWorkbookMap = WorkflowController.Instance.GetWorkbookMapForLayerMap(this);

                // Rebuild the Layer Details view model/Custom Task Pane if the layer's properties are updated in WWT.
                if (layerVersion > LayerDetails.Version)
                {
                    if (IsPropertyChangedFromCode)
                    {
                        LayerDetails.Version = layerVersion;
                    }
                    else
                    {
                        // Get the current properties of the layer from WWT.
                        Layer layerDetails = WWTManager.GetLayerDetails(LayerDetails.ID, LayerDetails.Group, true);

                        if (layerDetails != null)
                        {
                            LayerDetails = layerDetails;

                            // Update the received latest properties to the LayerMap.
                            this.UpdateLayerMapProperties(LayerDetails);

                            // Save the workbook map to the workbook which it belongs to.
                            this.SaveWorkbookMap();

                            if (parentWorkbookMap != null && parentWorkbookMap.SelectedLayerMap == this)
                            {
                                // This will update the custom task pane.
                                e.Result = true;
                            }
                        }
                    }
                }
                else if (layerVersion == -1)
                {
                    // In case if the layer is deleted or WWT is closed, layerVersion will be returned as -1.
                    // In case of Timeout, layers current version will be returned which is expected in case of no update to the layer properties.
                    // Setting not in sync will stop the notification as well.
                    IsNotInSync = true;

                    if (parentWorkbookMap != null && parentWorkbookMap.SelectedLayerMap == this)
                    {
                        if (MapType == LayerMapType.WWT)
                        {
                            // If current layer map is of type WWT and it is selected in layer dropdown, make sure the selection is removed.
                            parentWorkbookMap.SelectedLayerMap = null;
                        }

                        // This will update the custom task pane.
                        e.Result = true;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
            }
        }

        /// <summary>
        /// Handles the RunWorkerCompleted event of the BackgroundWorker instance.
        /// </summary>
        /// <param name="sender">Background Worker</param>
        /// <param name="e">BackgroundWorker event arguments</param>
        private void OnPropertyChangeNotifierCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                BackgroundWorker propertyChangeNotifier = sender as BackgroundWorker;

                if ((bool)e.Result && WorkflowController.Instance.GetWorkbookMapForLayerMap(this).Workbook == ThisAddIn.ExcelApplication.ActiveWorkbook)
                {
                    // Update the custom task pane.
                    Dispatcher.CurrentDispatcher.BeginInvoke(
                            new System.Action(delegate { WorkflowController.Instance.BuildAndBindLayerDetailsViewModel(true, false); }),
                            DispatcherPriority.Normal);
                }

                // When cancellationTokenSource is null, BackgroundWoker to be disposed.
                if (this.cancellationTokenSource == null)
                {
                    if (propertyChangeNotifier != null)
                    {
                        propertyChangeNotifier.CancelAsync();
                        propertyChangeNotifier.Dispose();
                    }
                }
                else if (propertyChangeNotifier != null && !propertyChangeNotifier.IsBusy && !propertyChangeNotifier.CancellationPending)
                {
                    // Again start listening to the notifications from WWT.
                    IsPropertyChangedFromCode = false;
                    propertyChangeNotifier.RunWorkerAsync();
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
            }
        }

        /// <summary>
        /// Removes Latitude and Longitude column mappings from the current layer mappings details..
        /// </summary>
        private void RemoveLatLongMappings()
        {
            if (MappedColumnType.IndexOf(ColumnType.Lat) > -1)
            {
                MappedColumnType[MappedColumnType.IndexOf(ColumnType.Lat)] = ColumnType.None;
            }

            if (MappedColumnType.IndexOf(ColumnType.Long) > -1)
            {
                MappedColumnType[MappedColumnType.IndexOf(ColumnType.Long)] = ColumnType.None;
            }
        }

        /// <summary>
        /// Removes RA and Dec column mappings from the current layer mappings details.
        /// </summary>
        private void RemoveRaDecMappings()
        {
            if (MappedColumnType.IndexOf(ColumnType.RA) > -1)
            {
                MappedColumnType[MappedColumnType.IndexOf(ColumnType.RA)] = ColumnType.None;
            }

            if (MappedColumnType.IndexOf(ColumnType.Dec) > -1)
            {
                MappedColumnType[MappedColumnType.IndexOf(ColumnType.Dec)] = ColumnType.None;
            }
        }

        /// <summary>
        /// Removes Altitude or Depth or Distance column mappings from the current layer mappings details, if they exists.
        /// </summary>
        private void RemoveDepthColumnMappings()
        {
            if (MappedColumnType.IndexOf(ColumnType.Alt) > -1)
            {
                MappedColumnType[MappedColumnType.IndexOf(ColumnType.Alt)] = ColumnType.None;
            }

            if (MappedColumnType.IndexOf(ColumnType.Depth) > -1)
            {
                MappedColumnType[MappedColumnType.IndexOf(ColumnType.Depth)] = ColumnType.None;
            }

            if (MappedColumnType.IndexOf(ColumnType.Distance) > -1)
            {
                MappedColumnType[MappedColumnType.IndexOf(ColumnType.Distance)] = ColumnType.None;
            }
        }

        #endregion
    }
}
