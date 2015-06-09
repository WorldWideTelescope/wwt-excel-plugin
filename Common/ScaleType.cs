//-----------------------------------------------------------------------
// <copyright file="ScaleType.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Enumeration of ScaleType.
    /// </summary>
    public enum ScaleType
    {
        /// <summary>
        /// ScaleType is linear.
        /// </summary>
        Linear,

        /// <summary>
        /// ScaleType is power.
        /// </summary>
        Power,

        /// <summary>
        /// ScaleType is Logarithmic.
        /// </summary>
        Log,

        /// <summary>
        /// ScaleType is constant.
        /// </summary>
        Constant,

        /// <summary>
        /// ScaleType is StellarMagnitude
        /// </summary>
        StellarMagnitude
    }
}
