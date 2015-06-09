//-----------------------------------------------------------------------
// <copyright file="LayerMapType.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Enumeration for identifying the type of layer.
    /// </summary>
    public enum LayerMapType
    {
        /// <summary>
        /// If the layer is Select One
        /// </summary>
        None,

        /// <summary>
        /// If the layer is present in only workbook.
        /// </summary>
        Local,

        /// <summary>
        /// If the layer is present only in WWT .
        /// </summary>
        WWT,

        /// <summary>
        /// If the layer is present in local and is created in WWT.
        /// </summary>
        LocalInWWT
    }
}
