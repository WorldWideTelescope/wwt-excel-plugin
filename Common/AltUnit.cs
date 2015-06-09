//-----------------------------------------------------------------------
// <copyright file="Layer.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Enumeration to identify the units for WWT.
    /// </summary>
    public enum AltUnit
    {
        /// <summary>
        /// If unit is in inches.
        /// </summary>
        Inches,

        /// <summary>
        /// If unit is in feet.
        /// </summary>
        Feet,

        /// <summary>
        /// If unit is in miles.
        /// </summary>
        Miles,

        /// <summary>
        /// If unit is in meters.
        /// </summary>
        Meters,

        /// <summary>
        /// If unit is in kilometers.
        /// </summary>
        Kilometers,

        /// <summary>
        /// If unit is in AstronomicalUnits.
        /// </summary>
        AstronomicalUnits,

        /// <summary>
        /// If unit is in LightYears.
        /// </summary>
        LightYears,

        /// <summary>
        /// If unit is in Parsecs.
        /// </summary>
        Parsecs,

        /// <summary>
        /// If unit is in MegaParsecs.
        /// </summary>
        MegaParsecs,

        /// <summary>
        /// Custom units.
        /// </summary>
        Custom
    }
}
