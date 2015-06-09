//-----------------------------------------------------------------------
// <copyright file="ViewpointMap.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Collections.ObjectModel;
using System.Runtime.Serialization;
using Microsoft.Office.Interop.Excel;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// This class stores the Viewpoints associated with the workbook.
    /// </summary>
    [DataContract]
    internal class ViewpointMap
    {
        /// <summary>
        /// Initializes a new instance of the ViewpointMap class.
        /// </summary>
        /// <param name="workbook">
        /// workbook object
        /// </param>
        internal ViewpointMap(Workbook workbook)
        {
            this.Workbook = workbook;
            this.SerializablePerspective = new ObservableCollection<Perspective>();
        }

        /// <summary>
        /// Gets or sets list of Viewpoints which will be serialized
        /// </summary>
        [DataMember]
        internal ObservableCollection<Perspective> SerializablePerspective { get; set; }

        /// <summary>
        /// Gets or sets the workbook in the current instance.
        /// </summary>
        internal Workbook Workbook { get; set; }
    }
}
