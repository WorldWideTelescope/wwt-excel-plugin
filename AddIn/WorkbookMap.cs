//-----------------------------------------------------------------------
// <copyright file="WorkbookMap.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Collections.Generic;
using System.Runtime.Serialization;
using Microsoft.Office.Interop.Excel;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// This class stores the the workbook and the associated layer with the workbook.
    /// This class is the core object model which needs to be kept in sync with excel and layer manager
    /// pane events so that it can be persisted at any point of time
    /// </summary>
    [DataContract]
    internal class WorkbookMap
    {
        /// <summary>
        /// Initializes a new instance of the WorkbookMap class.
        /// </summary>
        /// <param name="workbook">
        /// workbook object
        /// </param>
        internal WorkbookMap(Workbook workbook)
        {
            this.Workbook = workbook;
            this.SerializableLayerMaps = new List<LayerMap>();
            this.AllLayerMaps = new List<LayerMap>();
        }

        /// <summary>
        /// Gets or sets list of layer maps which will be serialized
        /// </summary>
        [DataMember]
        internal List<LayerMap> SerializableLayerMaps { get; set; }

        /// <summary>
        /// Gets or sets selected layer map which will be serialized
        /// </summary>
        [DataMember]
        internal LayerMap SerializableSelectedLayerMap { get; set; }

        /// <summary>
        /// Gets or sets the workbook in the current instance.
        /// </summary>
        internal Workbook Workbook { get; set; }

        /// <summary>
        /// Gets the layer map instance for valid local layers.
        /// </summary>
        internal List<LayerMap> LocalLayerMaps
        {
            get
            {
                List<LayerMap> maps = new List<LayerMap>();
                foreach (LayerMap item in AllLayerMaps)
                {
                    bool flag = false;
                    try
                    {
                        flag = item.IsValid && (item.MapType == Common.LayerMapType.LocalInWWT || item.MapType == Common.LayerMapType.Local);
                    }
                    catch (System.NullReferenceException)
                    {
                        if (ThisAddIn.ExcelApplication.ActiveWorkbook == null)
                        {
                            item.WorkbookReference = this.Workbook;
                            flag = item.IsValid && (item.MapType == Common.LayerMapType.LocalInWWT || item.MapType == Common.LayerMapType.Local);
                            item.WorkbookReference = null;
                        }
                    }
                    if (flag) 
                        maps.Add(item);
                }
                return maps;
            }
        }

        /// <summary>
        /// Gets only local in WWT layers.
        /// </summary>
        internal List<LayerMap> LocalInWWTLayerMaps
        {
            get
            {
                return AllLayerMaps.FindAll(item =>
                {
                    return item.IsValid && item.MapType == Common.LayerMapType.LocalInWWT;
                });
            }
        }

        /// <summary>
        /// Gets or sets all layer map instances including invalid layer map.
        /// </summary>
        internal List<LayerMap> AllLayerMaps { get; set; }

        /// <summary>
        /// Gets the layer map instance for WWT layers.
        /// </summary>
        internal List<LayerMap> WWTLayerMaps
        {
            get
            {
                return AllLayerMaps.FindAll(item =>
                {
                    return item.MapType == Common.LayerMapType.WWT;
                });
            }
        }

        /// <summary>
        /// Gets or sets the selected layer map based on SelectedLayerMapID
        /// Required for Custom task pane view model on workbook change
        /// </summary>
        internal LayerMap SelectedLayerMap { get; set; }

        /// <summary>
        /// This method gets called to manipulate the object before serialization occurs
        /// </summary>
        /// <param name="context">Streaming Context</param>
        [OnSerializing()]
        private void OnSerializingMethod(StreamingContext context)
        {
            SerializableLayerMaps.Clear();
            SerializableSelectedLayerMap = null;
            SerializableLayerMaps.AddRange(LocalLayerMaps);
            if (SelectedLayerMap != null && WorkflowController.IsLocalLayer(SelectedLayerMap.MapType))
            {
                if (SelectedLayerMap.IsValid)
                {
                    SerializableSelectedLayerMap = SelectedLayerMap;
                }
                else
                {
                    SelectedLayerMap = null;
                }
            }
        }
    }
}
