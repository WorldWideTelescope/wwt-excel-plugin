//-----------------------------------------------------------------------
// <copyright file="WorkbookExtensions.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// This class has all the extensions classes required for Excel.Workbook class.
    /// </summary>
    internal static class WorkbookExtensions
    {
        /// <summary>
        /// Get WorkbookMap from Workbook
        /// </summary>
        /// <param name="workbook">workbook instance</param>
        /// <returns>WorkbookMap instance</returns>
        internal static WorkbookMap GetWorkbookMap(this Workbook workbook)
        {
            // Initialize default
            var workbookMap = new WorkbookMap(workbook);
            if (workbook != null)
            {
                string content = workbook.GetCustomXmlPart(Common.Constants.XmlNamespace);
                if (!string.IsNullOrEmpty(content))
                {
                    workbookMap = workbookMap.Deserialize(content);

                    workbookMap.Workbook = workbook;
                    if (workbookMap.SerializableLayerMaps == null)
                    {
                        workbookMap.SerializableLayerMaps = new List<LayerMap>();
                    }

                    if (workbookMap.AllLayerMaps == null)
                    {
                        workbookMap.AllLayerMaps = new List<LayerMap>();
                        workbookMap.AllLayerMaps.AddRange(workbookMap.SerializableLayerMaps);
                    }

                    if (workbookMap.AllLayerMaps.Count > 0)
                    {
                        // Clean all invalid named ranges.
                        CleanLayerMap(workbook, workbookMap);

                        // Loop through all the layers to check if the range address column count is different from the mapped column count
                        // If so, reset the mapping as excel has undergone changes without the add-in
                        foreach (LayerMap localLayer in workbookMap.LocalLayerMaps)
                        {
                            if (localLayer.RangeName.RefersToRange != null && localLayer.RangeName.RefersToRange.EntireColumn.Count != localLayer.MappedColumnType.Count)
                            {
                                localLayer.SetAutoMap();
                                localLayer.SetLayerProperties();
                            }
                        }
                    }

                    // Check if layer being removed is selected layer
                    if (workbookMap.SerializableSelectedLayerMap != null)
                    {
                        // Check if SelectedLayerMap is found in All local layers
                        LayerMap layerMap = workbookMap.AllLayerMaps
                            .Where(item => !string.IsNullOrEmpty(item.RangeDisplayName) && item.RangeDisplayName.Equals(workbookMap.SerializableSelectedLayerMap.RangeDisplayName)).FirstOrDefault();

                        workbookMap.SelectedLayerMap = (layerMap != null) ? layerMap : null;
                    }
                }

                // Clean up and update Local in WWT Layers 
                workbookMap.LoadWWTLayers();
            }

            return workbookMap;
        }

        /// <summary>
        /// Get ViewpointMap from Workbook
        /// </summary>
        /// <param name="workbook">workbook instance</param>
        /// <returns>ViewpointMap instance</returns>
        internal static ViewpointMap GetViewpointMap(this Workbook workbook)
        {
            // Initialize default
            var viewpointMap = new ViewpointMap(workbook);
            if (workbook != null)
            {
                string content = workbook.GetCustomXmlPart(Common.Constants.ViewpointMapXmlNamespace);
                if (!string.IsNullOrEmpty(content))
                {
                    viewpointMap = viewpointMap.Deserialize(content);
                    viewpointMap.Workbook = workbook;
                    if (viewpointMap.SerializablePerspective == null)
                    {
                        viewpointMap.SerializablePerspective = new ObservableCollection<Perspective>();
                    }
                }
            }

            return viewpointMap;
        }

        /// <summary>
        /// Saved workbook map into Custom xml parts in workbook
        /// </summary>
        /// <param name="workbook">workbook instance</param>
        /// <param name="workBookMaps">List of all workbook maps</param>
        internal static void SaveWorkbookMap(this Workbook workbook, List<WorkbookMap> workBookMaps)
        {
            var workbookMap = workBookMaps.Find(item => item.Workbook == workbook);
            if (workbookMap != null)
            {
                // Save only local layers
                if (workbookMap.LocalLayerMaps != null)
                {
                    lock (WorkflowController.LockObject)
                    {
                        string content = workbookMap.Serialize();
                        if (!string.IsNullOrEmpty(content))
                        {
                            workbook.AddCustomXmlPart(content, Common.Constants.XmlNamespace);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// This function is used to remove all layer which are not valid.
        ///     1. Layer map which does not have a valid named range
        /// </summary>
        /// <param name="workbook">
        /// workbook instance
        /// </param>
        /// <param name="workbookMap">
        /// workbook map instance
        /// </param>
        private static void CleanLayerMap(Workbook workbook, WorkbookMap workbookMap)
        {
            // Clean up Invalid Ranges.
            foreach (Name namedRange in workbook.Names)
            {
                // The named ranges are WWT ranges and are invalid
                if (namedRange.IsWWTRange() && !namedRange.IsValid())
                {
                    string name = namedRange.Name;
                    try
                    {
                        // Delete the range name for all invalid layers
                        workbookMap.AllLayerMaps.ForEach(layer =>
                        {
                            if (layer.RangeDisplayName.Equals(name))
                            {
                                layer.RangeName.Delete();
                            }
                        });
                    }
                    catch (COMException ex)
                    {
                        Logger.LogException(ex);
                    }

                    // Clean up such layers on load itself
                    workbookMap.AllLayerMaps.RemoveAll(layer =>
                    {
                        return layer.RangeDisplayName.Equals(name);
                    });
                }
            }
        }
    }
}
