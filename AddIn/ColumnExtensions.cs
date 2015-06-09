//-----------------------------------------------------------------------
// <copyright file="ColumnExtensions.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Extension methods on Column
    /// </summary>
    internal static class ColumnExtensions
    {
        #region internal methods

        /// <summary>
        /// Checks if the column is depth type of column
        /// </summary>
        /// <param name="column">column instance</param>
        /// <returns>true or false</returns>
        internal static bool IsDepthColumn(this Column column)
        {
            bool isDepthColumn = false;
            if (column != null)
            {
                if (column.ColType == ColumnType.Depth || column.ColType == ColumnType.Alt || column.ColType == ColumnType.Distance)
                {
                    isDepthColumn = true;
                }
            }

            return isDepthColumn;
        }

        /// <summary>
        /// Checks if the column is XYZ column.
        /// </summary>
        /// <param name="column">Column instance</param>
        /// <returns>True in case of XYZ column. Otherwise false.</returns>
        internal static bool IsXYZColumn(this Column column)
        {
            bool xyzColumn = false;

            if (column != null)
            {
                if (column.ColType == ColumnType.X ||
                        column.ColType == ColumnType.Y ||
                        column.ColType == ColumnType.Z ||
                        column.ColType == ColumnType.ReverseX ||
                        column.ColType == ColumnType.ReverseY ||
                        column.ColType == ColumnType.ReverseZ)
                {
                    xyzColumn = true;
                }
            }

            return xyzColumn;
        }

        /// <summary>
        /// Checks if the column is RA type of column
        /// </summary>
        /// <param name="column">column instance</param>
        /// <returns>true or false</returns>
        internal static bool IsRAColumn(this Column column)
        {
            return (column != null && column.ColType == ColumnType.RA);
        }

        /// <summary>
        /// Create a column object that represents the default column.
        /// This is typically captioned "Select One".
        /// </summary>
        /// <returns>Column object that represents the default column</returns>
        internal static Column GetDefaultColumn()
        {
            return new Column(ColumnType.None, Properties.Resources.SelectOneString, new Collection<string>());
        }

        /// <summary>
        /// Populates the column list
        /// </summary>
        /// <returns>List of columns</returns>
        internal static Collection<Column> PopulateColumnList()
        {
            return new Collection<Column>
            {
                GetDefaultColumn(),
                GetLatColumn(),
                GetLongColumn(),
                GetStartDateColumn(),
                GetEndDateColumn(),
                GetDepthColumn(),
                GetAltColumn(),
                GetDistanceColumn(),
                GetMagColumn(),
                GetGeometryColumn(),
                GetColorColumn(),
                GetRAColumn(),
                GetDECColumn(),
                GetXColumn(),
                GetYColumn(),
                GetZColumn(),
                GetReverseXColumn(),
                GetReverseYColumn(),
                GetReverseZColumn() 
            };
        }

        /// <summary>
        /// Create a column object that represents the "Depth" column
        /// </summary>
        /// <returns>Column object that represents the "Depth" column</returns>
        internal static Column GetDepthColumn()
        {
            return new Column(ColumnType.Depth, Properties.Resources.DepthDisplayValue, GetComparisionValuesForDepth());
        }

        /// <summary>
        /// Create a column object that represents the "Altitude" column
        /// </summary>
        /// <returns>Column object that represents the "Altitude" column</returns>
        internal static Column GetAltColumn()
        {
            return new Column(ColumnType.Alt, Properties.Resources.AltitudeDisplayValue, GetComparisionValuesForAlt());
        }

        /// <summary>
        /// Create a column object that represents the "Distance" column
        /// </summary>
        /// <returns>Column object that represents the "Distance" column</returns>
        internal static Column GetDistanceColumn()
        {
            return new Column(ColumnType.Distance, Properties.Resources.DistanceDisplayValue, GetComparisionValuesForDistance());
        }

        #endregion internal methods

        #region private methods

        /// <summary>
        /// Create a column object that represents the "Latitude" column
        /// </summary>
        /// <returns>Column object that represents the "Latitude" column</returns>
        private static Column GetLatColumn()
        {
            return new Column(ColumnType.Lat, Properties.Resources.LatitudeDisplayValue, GetComparisonValuesForLat());
        }

        /// <summary>
        /// Create a column object that represents the "Longitude" column
        /// </summary>
        /// <returns>Column object that represents the "Longitude" column</returns>
        private static Column GetLongColumn()
        {
            return new Column(ColumnType.Long, Properties.Resources.LongitudeDisplayValue, GetComparisonValuesForLong());
        }

        /// <summary>
        /// Create a column object that represents the "StartDate" column
        /// </summary>
        /// <returns>Column object that represents the "StartDate" column</returns>
        private static Column GetStartDateColumn()
        {
            return new Column(ColumnType.StartDate, Properties.Resources.StartDateDisplayValue, GetComparisionValuesForStartDate());
        }

        /// <summary>
        /// Create a column object that represents the "EndDate" column
        /// </summary>
        /// <returns>Column object that represents the "EndDate" column</returns>
        private static Column GetEndDateColumn()
        {
            return new Column(ColumnType.EndDate, Properties.Resources.EndDateDisplayValue, GetComparisionValuesForEndDate());
        }

        /// <summary>
        /// Create a column object that represents the "Magnitude" column
        /// </summary>
        /// <returns>Column object that represents the "Magnitude" column</returns>
        private static Column GetMagColumn()
        {
            return new Column(ColumnType.Mag, Properties.Resources.MagnitudeDisplayValue, GetComparisionValuesForMag());
        }

        /// <summary>
        /// Create a column object that represents the "Geometry" column
        /// </summary>
        /// <returns>Column object that represents the "Geometry" column</returns>
        private static Column GetGeometryColumn()
        {
            return new Column(ColumnType.Geo, Properties.Resources.GeometryDisplayValue, GetComparisionValuesForGeo());
        }

        /// <summary>
        /// Create a column object that represents the "Color" column
        /// </summary>
        /// <returns>Column object that represents the "Color" column</returns>
        private static Column GetColorColumn()
        {
            return new Column(ColumnType.Color, Properties.Resources.ColorDisplayValue, GetComparisionValuesForColor());
        }

        /// <summary>
        /// Create a column object that represents the "RA" column
        /// </summary>
        /// <returns>Column object that represents the "RA" column</returns>
        private static Column GetRAColumn()
        {
            return new Column(ColumnType.RA, Properties.Resources.RADisplayValue, GetComparisonValuesForRA());
        }

        /// <summary>
        /// Create a column object that represents the "DEC" column
        /// </summary>
        /// <returns>Column object that represents the "DEC" column</returns>
        private static Column GetDECColumn()
        {
            return new Column(ColumnType.Dec, Properties.Resources.DecDisplayValue, GetComparisonValuesForDEC());
        }

        /// <summary>
        /// Create a column object that represents the "X" column
        /// </summary>
        /// <returns>Column object that represents the "X" column</returns>
        private static Column GetXColumn()
        {
            return new Column(ColumnType.X, Properties.Resources.XDisplayValue, GetComparisionValuesForX());
        }

        /// <summary>
        /// Create a column object that represents the "Y" column
        /// </summary>
        /// <returns>Column object that represents the "Y" column</returns>
        private static Column GetYColumn()
        {
            return new Column(ColumnType.Y, Properties.Resources.YDisplayValue, GetComparisionValuesForY());
        }

        /// <summary>
        /// Create a column object that represents the "Z" column
        /// </summary>
        /// <returns>Column object that represents the "Z" column</returns>
        private static Column GetZColumn()
        {
            return new Column(ColumnType.Z, Properties.Resources.ZDisplayValue, GetComparisionValuesForZ());
        }

        /// <summary>
        /// Create a column object that represents the "-X" column
        /// </summary>
        /// <returns>Column object that represents the "-X" column</returns>
        private static Column GetReverseXColumn()
        {
            return new Column(ColumnType.ReverseX, Properties.Resources.XReverseDisplayValue, GetComparisionValuesForReverseX());
        }

        /// <summary>
        /// Create a column object that represents the "-Y" column
        /// </summary>
        /// <returns>Column object that represents the "-Y" column</returns>
        private static Column GetReverseYColumn()
        {
            return new Column(ColumnType.ReverseY, Properties.Resources.YReverseDisplayValue, GetComparisionValuesForReverseY());
        }

        /// <summary>
        /// Create a column object that represents the "-Z" column
        /// </summary>
        /// <returns>Column object that represents the "-Z" column</returns>
        private static Column GetReverseZColumn()
        {
            return new Column(ColumnType.ReverseZ, Properties.Resources.ZReverseDisplayValue, GetComparisionValuesForReverseZ());
        }

        /// <summary>
        /// Build a collection of strings to match the "Latitude" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "Latitude" column heading</returns>
        private static Collection<string> GetComparisonValuesForLat()
        {
            return new Collection<string>
            {
                Properties.Resources.LatitudeDisplayValue,
                Properties.Resources.LatitudeComparison1
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "RA" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "RA" column heading</returns>
        private static Collection<string> GetComparisonValuesForRA()
        {
            return new Collection<string>
            {
                Properties.Resources.RAComparisonValue1,
                Properties.Resources.RAComparisonValue2,
                Properties.Resources.RADisplayValue
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "DEC" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "DEC" column heading</returns>
        private static Collection<string> GetComparisonValuesForDEC()
        {
            return new Collection<string>
            {
                Properties.Resources.DecComparisonValue1,
                Properties.Resources.DecDisplayValue
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "Longitude" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "Longitude" column heading</returns>
        private static Collection<string> GetComparisonValuesForLong()
        {
            return new Collection<string>
            {
                Properties.Resources.LongitudeDisplayValue,
                Properties.Resources.LongitudeComparison1
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "StartDate" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "StartDate" column heading</returns>
        private static Collection<string> GetComparisionValuesForStartDate()
        {
            return new Collection<string>
            {
                Properties.Resources.StartDateDisplayValue,
                Properties.Resources.StartDateComparision1,
                Properties.Resources.StartDateComparision2,
                Properties.Resources.StartDateComparision3,
                Properties.Resources.StartDateComparision4,
                Properties.Resources.StartDateComparision5,
                Properties.Resources.StartDateComparision6,
                Properties.Resources.StartDateComparision7,
                Properties.Resources.StartDateComparision8,
                Properties.Resources.StartDateComparision9,
                Properties.Resources.StartDateComparision10,
                Properties.Resources.StartDateComparision11,
                Properties.Resources.StartDateComparision12
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "EndDate" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "EndDate" column heading</returns>
        private static Collection<string> GetComparisionValuesForEndDate()
        {
            return new Collection<string>
            {
                Properties.Resources.EndDateDisplayValue,
                Properties.Resources.EndDateComparision1,
                Properties.Resources.EndDateComparision2,
                Properties.Resources.EndDateComparision3,
                Properties.Resources.EndDateComparision4
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "Depth" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "Depth" column heading</returns>
        private static Collection<string> GetComparisionValuesForDepth()
        {
            return new Collection<string>
            {
                Properties.Resources.DepthDisplayValue
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "Altitude" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "Altitude" column heading</returns>
        private static Collection<string> GetComparisionValuesForAlt()
        {
            return new Collection<string>
            {
                Properties.Resources.AltitudeDisplayValue,
                Properties.Resources.AltitudeComparision3,
                Properties.Resources.AltitudeComparision1,
                Properties.Resources.AltitudeComparision2
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "Distance" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "Distance" column heading</returns>
        private static Collection<string> GetComparisionValuesForDistance()
        {
            return new Collection<string>
            {
                Properties.Resources.DistanceDisplayValue,
                Properties.Resources.DistanceComparision2,
                Properties.Resources.DistanceComparision1
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "Magnitude" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "Magnitude" column heading</returns>
        private static Collection<string> GetComparisionValuesForMag()
        {
            return new Collection<string>
            {
                Properties.Resources.MagnitudeDisplayValue,
                Properties.Resources.MagnitudeComparision1
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "Geometry" display value
        /// </summary>
        /// <returns>Collection of strings matching the "Geometry" column heading</returns>
        private static Collection<string> GetComparisionValuesForGeo()
        {
            return new Collection<string>
            {
                Properties.Resources.GeometryDisplayValue,
                Properties.Resources.GeometryComparision1
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "Color" display value
        /// </summary>
        /// <returns>Collection of strings matching the "Color" column heading</returns>
        private static Collection<string> GetComparisionValuesForColor()
        {
            return new Collection<string>
            {
                Properties.Resources.ColorDisplayValue
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "X" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "X" column heading</returns>
        private static Collection<string> GetComparisionValuesForX()
        {
            return new Collection<string>
            {
                Properties.Resources.XDisplayValue
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "Y" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "Y" column heading</returns>
        private static Collection<string> GetComparisionValuesForY()
        {
            return new Collection<string>
            {
                Properties.Resources.YDisplayValue
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "Z" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "Z" column heading</returns>
        private static Collection<string> GetComparisionValuesForZ()
        {
            return new Collection<string>
            {
                Properties.Resources.ZDisplayValue
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "-X" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "-X" column heading</returns>
        private static Collection<string> GetComparisionValuesForReverseX()
        {
            return new Collection<string>
            {
                Properties.Resources.XReverseDisplayValue,
                Properties.Resources.XReverseComparision1
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "-Y" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "-Y" column heading</returns>
        private static Collection<string> GetComparisionValuesForReverseY()
        {
            return new Collection<string>
            {
                Properties.Resources.YReverseDisplayValue,
                Properties.Resources.YReverseComparision1
            };
        }

        /// <summary>
        /// Build a collection of strings to match the "-Z" column heading
        /// </summary>
        /// <returns>Collection of strings matching the "-Z" column heading</returns>
        private static Collection<string> GetComparisionValuesForReverseZ()
        {
            return new Collection<string>
            {
                Properties.Resources.ZReverseDisplayValue,
                Properties.Resources.ZReverseComparision1
            };
        }

        #endregion private methods
    }
}
