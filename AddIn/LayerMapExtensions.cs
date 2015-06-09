//-----------------------------------------------------------------------
// <copyright file="LayerMapExtensions.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Extension class for Layer map
    /// </summary>
    internal static class LayerMapExtensions
    {
        /// <summary>
        /// Updates the layer map properties with the layer values
        /// </summary>
        /// <param name="layerMap">Layer map value</param>
        /// <param name="layer">Layer details</param>
        /// <returns>Updated layer map</returns>
        internal static LayerMap UpdateLayerMapProperties(this LayerMap layerMap, Layer layer)
        {
            if (layerMap != null && layer != null)
            {
                lock (WorkflowController.LockObject)
                {
                    layerMap.LayerDetails = layer;
                    layerMap.ColumnsList = ColumnExtensions.PopulateColumnList();
                    layerMap.SetMappedColumnType();
                }
            }
            return layerMap;
        }

        /// <summary>
        /// Update the mapped column types based on the Selected Layer map.
        /// </summary>
        /// <param name="selectedLayerMap">
        /// Selected Layer map.
        /// </param>
        internal static void UpdateMappedColumns(this LayerMap selectedLayerMap)
        {
            if (selectedLayerMap != null && selectedLayerMap.LayerDetails.Group != null)
            {
                Collection<ColumnType> mappedColTypes = new Collection<ColumnType>();
                if (selectedLayerMap.LayerDetails.Group.IsPlanet())
                {
                    selectedLayerMap.MappedColumnType.ToList().ForEach(columnType =>
                    {
                        if (columnType == ColumnType.RA || columnType == ColumnType.Dec)
                        {
                            mappedColTypes.Add(ColumnType.None);
                        }
                        else
                        {
                            mappedColTypes.Add(columnType);
                        }
                    });
                }
                else
                {
                    selectedLayerMap.MappedColumnType.ToList().ForEach(columnType =>
                    {
                        if (columnType == ColumnType.Lat || columnType == ColumnType.Long)
                        {
                            mappedColTypes.Add(ColumnType.None);
                        }
                        else
                        {
                            mappedColTypes.Add(columnType);
                        }
                    });
                }

                selectedLayerMap.MappedColumnType.Clear();
                selectedLayerMap.MappedColumnType = mappedColTypes;
            }
        }

        /// <summary>
        /// Saved workbook map into Custom xml parts in workbook
        /// </summary>
        /// <param name="layerMap">LayerMap instance</param>
        internal static void SaveWorkbookMap(this LayerMap layerMap)
        {
            if (layerMap != null)
            {
                var workbookMap = WorkflowController.Instance.GetWorkbookMapForLayerMap(layerMap);
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
                                workbookMap.Workbook.AddCustomXmlPart(content, Common.Constants.XmlNamespace);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Checks if the layer is already created in WWT. 
        /// </summary>
        /// <param name="selectedLayerMap">Selected layer map</param>
        /// <returns>True if the layer is created in WWT;Otherwise false.</returns>
        internal static bool IsLayerCreated(this LayerMap selectedLayerMap)
        {
            return selectedLayerMap != null && string.IsNullOrEmpty(selectedLayerMap.LayerDetails.ID);
        }

        /// <summary>
        /// This function is used to check if we can update WWT or not.
        /// </summary>
        /// <param name="selectedlayer">
        /// LayerMap in focus.
        /// </param>
        /// <returns>
        /// True, if we can update WWT;Otherwise false.
        /// </returns>
        internal static bool CanUpdateWWT(this LayerMap selectedlayer)
        {
            return (selectedlayer != null && selectedlayer.MapType == LayerMapType.LocalInWWT && !selectedlayer.IsNotInSync) || selectedlayer.MapType == LayerMapType.WWT;
        }

        /// <summary>
        /// Checks if the given layer is XYZ layer or not.
        /// If X/ReverseX or Y/ReverseY or Z/ReverseZ is mapped and Lat/Long or RA/DEC is not mapped, then the layer will considered as XYZ layer.
        /// </summary>
        /// <param name="selectedLayerMap">Selected layer</param>
        /// <returns>True, if RA/DEC/Lat/Long are not mapped and X or Y or Z is/are mapped. False, otherwise.</returns>
        internal static bool IsXYZLayer(this LayerMap selectedLayerMap)
        {
            bool xyzLayer = false;

            if (selectedLayerMap.MappedColumnType != null)
            {
                xyzLayer = (selectedLayerMap.MappedColumnType.Contains(ColumnType.X) || selectedLayerMap.MappedColumnType.Contains(ColumnType.Y) || selectedLayerMap.MappedColumnType.Contains(ColumnType.Z)
                   || selectedLayerMap.MappedColumnType.Contains(ColumnType.ReverseX) || selectedLayerMap.MappedColumnType.Contains(ColumnType.ReverseY)
                   || selectedLayerMap.MappedColumnType.Contains(ColumnType.ReverseZ)) && !(selectedLayerMap.MappedColumnType.Contains(ColumnType.Lat) || selectedLayerMap.MappedColumnType.Contains(ColumnType.Long)
                   || selectedLayerMap.MappedColumnType.Contains(ColumnType.RA) || selectedLayerMap.MappedColumnType.Contains(ColumnType.Dec));
            }

            return xyzLayer;
        }

        /// <summary>
        /// Checks if the given layer is mapped with XY columns or not.
        /// If X/ReverseX and Y/ReverseY are mapped and Lat/Long/RA/DEC are not mapped (X and Y are mandatory), layer will be considered as XY mapped.
        /// </summary>
        /// <param name="selectedLayerMap">Selected layer map</param>
        /// <returns>True if RA/DEC/Lat/Long are not mapped and X and Y are mapped. False, otherwise.</returns>
        internal static bool IsXYMappedLayer(this LayerMap selectedLayerMap)
        {
            return (selectedLayerMap.MappedColumnType.Contains(ColumnType.X) || selectedLayerMap.MappedColumnType.Contains(ColumnType.ReverseX))
               && (selectedLayerMap.MappedColumnType.Contains(ColumnType.Y) || selectedLayerMap.MappedColumnType.Contains(ColumnType.ReverseY))
               && !(selectedLayerMap.MappedColumnType.Contains(ColumnType.Lat) || selectedLayerMap.MappedColumnType.Contains(ColumnType.Long)
               || selectedLayerMap.MappedColumnType.Contains(ColumnType.RA) || selectedLayerMap.MappedColumnType.Contains(ColumnType.Dec));
        }

        /// <summary>
        /// Builds group collection with the current layer map and groups
        /// </summary>
        /// <param name="layerMap">Selected layer map</param>
        /// <param name="groups">Collection of groups</param>
        /// <returns>Collection of groups with reference frame and layers</returns>
        internal static List<GroupChildren> BuildGroupCollection(this LayerMap layerMap, List<GroupChildren> groups)
        {
            if (layerMap != null && layerMap.LayerDetails.Group != null && groups != null)
            {
                if (layerMap.LayerDetails.Group.Parent != null)
                {
                    GroupChildren groupItem = groups.Where(groupValue => groupValue.Group != null && groupValue.Group.Equals(layerMap.LayerDetails.Group.Parent)).FirstOrDefault();
                    if (groupItem != null)
                    {
                        GroupChildren groupChildNode = groupItem.Children.Where(childNodes => childNodes.Group.Equals(layerMap.LayerDetails.Group)).FirstOrDefault();
                        if (groupChildNode != null)
                        {
                            layerMap.Name = LayerDetailsViewModel.GetLayerNameOnMapType(layerMap, layerMap.LayerDetails.Name);
                            groupChildNode.AllChildren.Add(layerMap);
                        }
                        else
                        {
                            GroupChildren childNode = AddChildNodesToGroup(layerMap);
                            groupItem.Children.Add(childNode);
                            groupItem.AllChildren.Add(childNode);
                        }
                    }
                    else
                    {
                        AddLayerNode(groups, layerMap);
                    }
                }
                else
                {
                    AddLayerNode(groups, layerMap);
                }
            }
            return groups;
        }

        /// <summary>
        /// This function is used to update header details in object models.
        /// </summary>
        /// <param name="selectedlayer">
        /// Updated layer.
        /// </param>
        /// <param name="selectedRange">
        /// Updated range.
        /// </param>
        internal static void UpdateHeaderProperties(this LayerMap selectedlayer, Range selectedRange)
        {
            if (selectedlayer != null && selectedRange != null)
            {
                // Update the address of the selected layer.
                selectedlayer.RangeAddress = selectedRange.Address;

                // Update Header Data.
                selectedlayer.HeaderRowData = selectedRange.GetHeader();

                // Header Change
                // 1. AutoMap the columns 
                selectedlayer.SetAutoMap();

                // 2. Set layer properties dependent on mapping. 
                selectedlayer.SetLayerProperties();
            }
        }

        /// <summary>
        /// Gets the look at value from the selected layer
        /// </summary>
        /// <param name="layerMap">Selected layer map</param>
        /// <returns>Look at value from layer</returns>
        internal static string GetLookAt(this LayerMap layerMap)
        {
            string lookAt = Common.Constants.EarthLookAt;
            if (layerMap != null)
            {
                var referenceFramePath = layerMap.LayerDetails.Group.Path;
                if (referenceFramePath.StartsWith(Common.Constants.SkyFramePath, StringComparison.OrdinalIgnoreCase))
                {
                    lookAt = Common.Constants.SkyLookAt;
                }
                else if (!referenceFramePath.StartsWith(Common.Constants.EarthFramePath, StringComparison.OrdinalIgnoreCase))
                {
                    lookAt = Common.Constants.SolarSystemLookAt;
                }
            }

            return lookAt;
        }

        /// <summary>
        /// Add child node to existing group
        /// </summary>
        /// <param name="layerMap">Selected layer map</param>
        /// <returns>Child node for group children</returns>
        private static GroupChildren AddChildNodesToGroup(LayerMap layerMap)
        {
            GroupChildren childNode = new GroupChildren();
            if (layerMap != null)
            {
                childNode.Group = layerMap.LayerDetails.Group;
                childNode.Name = layerMap.LayerDetails.Group.Name;
                childNode.Layers.Add(layerMap.LayerDetails);
                layerMap.Name = LayerDetailsViewModel.GetLayerNameOnMapType(layerMap, layerMap.LayerDetails.Name);
                childNode.AllChildren.Add(layerMap);
                childNode.IsDeleted = layerMap.LayerDetails.Group.IsDeleted;
            }
            return childNode;
        }

        /// <summary>
        /// Creates nodes and add layer to the existing node
        /// </summary>
        /// <param name="groups">Group collection with layer and reference frame groups</param>
        /// <param name="layerMap">Selected layer map details</param>
        private static void AddLayerNode(List<GroupChildren> groups, LayerMap layerMap)
        {
            if (layerMap != null && groups != null)
            {
                GroupChildren parent = CreateNode(layerMap.LayerDetails.Group, groups, layerMap);
                layerMap.Name = LayerDetailsViewModel.GetLayerNameOnMapType(layerMap, layerMap.LayerDetails.Name);

                // Add Layer to the parent group.
                parent.AllChildren.Add(layerMap);
            }
        }

        /// <summary>
        /// Creates node with existing hierarchy
        /// </summary>
        /// <param name="group">Group in the hierarchy</param>
        /// <param name="groups">Group collection</param>
        /// <param name="layerMap">Selected layer map details</param>
        /// <returns>Last created group children</returns>
        private static GroupChildren CreateNode(Group group, List<GroupChildren> groups, LayerMap layerMap)
        {
            GroupChildren groupChildren = null;
            if (group.Parent != null)
            {
                GroupChildren parent = CreateNode(group.Parent, groups, layerMap);

                groupChildren = parent.Children.Where(child => child.Group != null && child.Group.Equals(group)).FirstOrDefault();
                if (groupChildren == null)
                {
                    groupChildren = new GroupChildren();
                    groupChildren.Group = group;
                    groupChildren.Name = group.Name;
                    groupChildren.IsDeleted = group.IsDeleted;

                    parent.AllChildren.Add(groupChildren);
                    parent.Children.Add(groupChildren);
                }
            }
            else
            {
                groupChildren = groups.Where(groupValue => groupValue.Group != null && groupValue.Group.Equals(group)).FirstOrDefault();
                if (groupChildren == null)
                {
                    groupChildren = new GroupChildren();
                    groupChildren.Group = group;
                    groupChildren.Name = group.Name;
                    groupChildren.IsDeleted = group.IsDeleted;
                    groups.Add(groupChildren);
                }
            }

            return groupChildren;
        }
    }
}
