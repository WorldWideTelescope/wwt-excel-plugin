//-----------------------------------------------------------------------
// <copyright file="ColumnType.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Enumeration for Column types.
    /// </summary>
    public enum ColumnType
    {
        /// <summary>
        /// Default type
        /// </summary>
        None,

        /// <summary>
        /// If column is Latitude.
        /// </summary>
        Lat,

        /// <summary>
        /// If column is Longitude
        /// </summary>
        Long,

        /// <summary>
        /// If column is Start date time
        /// </summary>
        StartDate,

        /// <summary>
        /// If column is End date time
        /// </summary>
        EndDate,

        /// <summary>
        /// If column is Depth
        /// </summary>
        Depth,

        /// <summary>
        /// If column is Altitude
        /// </summary>
        Alt,

        /// <summary>
        /// If column is Distance
        /// </summary>
        Distance,

        /// <summary>
        /// If column is Magnitude
        /// </summary>
        Mag,

        /// <summary>
        /// If column is Geography
        /// </summary>
        Geo,

        /// <summary>
        /// If column is Color
        /// </summary>
        Color,

        /// <summary>
        /// If column is RA
        /// </summary>
        RA,

        /// <summary>
        /// If column is Dec
        /// </summary>
        Dec,

        /// <summary>
        /// If column is X
        /// </summary>
        X,

        /// <summary>
        /// If column is Y
        /// </summary>
        Y,

        /// <summary>
        /// If column is Z
        /// </summary>
        Z,

        /// <summary>
        /// If column is -X
        /// </summary>
        ReverseX,

        /// <summary>
        /// If column is -Y
        /// </summary>
        ReverseY,

        /// <summary>
        /// If column is -Z
        /// </summary>
        ReverseZ,
    }
}
