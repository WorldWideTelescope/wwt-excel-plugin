//-----------------------------------------------------------------------
// <copyright file="WorkbookMapExtensions.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Text;
using System.Xml;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Extension methods on WorkbookMap
    /// </summary>
    internal static class WorkbookMapExtensions
    {
        /// <summary>
        /// Extension method to serialize the workbook map
        /// </summary>
        /// <param name="workbookMap">workbookMap object</param>
        /// <returns>serialized string</returns>
        internal static string Serialize(this WorkbookMap workbookMap)
        {
            StringBuilder serializedString = new StringBuilder();
            if (workbookMap != null)
            {
                try
                {
                    using (var writer = XmlWriter.Create(serializedString))
                    {
                        var serializer = new DataContractSerializer(typeof(WorkbookMap), Common.Constants.XmlRootName, Common.Constants.XmlNamespace);
                        if (writer != null)
                        {
                            serializer.WriteObject(writer, workbookMap);
                        }
                    }
                }
                catch (ArgumentNullException ex)
                {
                    Logger.LogException(ex);
                }
                catch (InvalidDataContractException ex)
                {
                    Logger.LogException(ex);
                }
                catch (SerializationException ex)
                {
                    Logger.LogException(ex);
                }
            }

            return serializedString.ToString();
        }

        /// <summary>
        /// Extension method to de-serialize the string into a workbook map
        /// </summary>
        /// <param name="workbookMap">workbookMap object</param>
        /// <param name="xmlContent">xml content</param>
        /// <returns>populated workbookMap object</returns>
        internal static WorkbookMap Deserialize(this WorkbookMap workbookMap, string xmlContent)
        {
            if (workbookMap != null)
            {
                using (var stringReader = new StringReader(xmlContent))
                {
                    try
                    {
                        var reader = XmlReader.Create(stringReader);
                        {
                            var serializer = new DataContractSerializer(typeof(WorkbookMap), Common.Constants.XmlRootName, Common.Constants.XmlNamespace);
                            workbookMap = (WorkbookMap)serializer.ReadObject(reader, true);
                        }
                    }
                    catch (ArgumentNullException ex)
                    {
                        Logger.LogException(ex);
                    }
                    catch (SerializationException ex)
                    {
                        Logger.LogException(ex);
                    }
                }
            }

            return workbookMap;
        }

        /// <summary>
        /// This function is used to check if the layer with the specified ID is present in the workbook.
        /// </summary>
        /// <param name="workbookMap">
        /// Parent container.
        /// </param>
        /// <param name="layerID">
        /// Id of the Layer in focus.
        /// </param>
        /// <returns>
        /// True if the LayerId is part of workbook collection;Otherwise false.
        /// </returns>
        internal static bool Exists(this WorkbookMap workbookMap, string layerID)
        {
            return workbookMap != null && workbookMap.AllLayerMaps.Exists(
                            layer =>
                            {
                                return string.CompareOrdinal(layerID, layer.LayerDetails.ID) == 0;
                            });
        }

        /// <summary>
        /// This function is used to Load all WWT layers on the load of a workbook.
        /// This function performs the following operations:
        ///     1. Remove all WWT only layers.
        ///     2. On load of a workbook, no layers are connected with WWT.
        ///     3. Insert and update deleted WWT Layers.
        /// </summary>
        /// <param name="workbookMap">
        /// Parent container.
        /// </param>
        internal static void LoadWWTLayers(this WorkbookMap workbookMap)
        {
            if (workbookMap != null)
            {
                LayerMap selectedLayerMap = workbookMap.SelectedLayerMap;

                // 1. Remove all WWT only layers.
                workbookMap.RemoveWWTLayers();

                ICollection<Group> groups = WWTManager.GetAllWWTGroups(true);

                // 2. On load of a workbook, no layers are connected with WWT.
                workbookMap.LocalInWWTLayerMaps.ForEach(
                    layer =>
                    {
                        // 2.a. Set IsNotInSync.
                        layer.IsNotInSync = true;

                        // 2.b. Update all group references.
                        Group group = SearchGroup(layer.LayerDetails.Group.Name, layer.LayerDetails.Group.Path, groups);
                        if (group != null)
                        {
                            layer.LayerDetails.Group = group;
                        }
                        else
                        {
                            layer.LayerDetails.Group.IsDeleted = true;
                        }
                    });

                // 3. Synchronize WWT Layers
                workbookMap.SyncWWTLayers(groups);

                // 4. Set the Selected layer map to null if the Selected WWT layer was deleted in WWT. 
                //  If the selected layer is WWT layer/is localInWWT which is not deleted, this will be selected while 
                if (selectedLayerMap != null && selectedLayerMap.MapType == LayerMapType.WWT && !workbookMap.Exists(selectedLayerMap.LayerDetails.ID))
                {
                    workbookMap.SelectedLayerMap = null;
                }
            }
        }

        /// <summary>
        /// This function is used to refresh the layer drop down with the latest layers from WWT.
        /// This function performs the following operations:
        ///     1. Remove all WWT only layers.
        ///     2. Insert and update deleted WWT Layers.
        ///     3. Set the Selected layer map to null if the Selected WWT layer was deleted in WWT. 
        /// </summary>
        /// <param name="workbookMap">
        /// Parent container.
        /// </param>
        internal static void RefreshLayers(this WorkbookMap workbookMap)
        {
            if (workbookMap != null)
            {
                LayerMap selectedLayerMap = workbookMap.SelectedLayerMap;
                ICollection<Group> groups = WWTManager.GetAllWWTGroups(true);

                // 1. Remove all WWT only layers.
                workbookMap.RemoveWWTLayers();

                // 2. Insert and update deleted WWT Layers
                workbookMap.SyncWWTLayers(groups);

                // 3. Set the Selected layer map to null if the Selected WWT layer was deleted in WWT. 
                //  If the selected layer is WWT layer/is localInWWT which is not deleted, this will be selected while 
                if (selectedLayerMap != null && selectedLayerMap.MapType == LayerMapType.WWT && !workbookMap.Exists(selectedLayerMap.LayerDetails.ID))
                {
                    workbookMap.SelectedLayerMap = null;
                }
            }
        }

        /// <summary>
        /// This function is used to Cleanup WWT layers from layer drop down when WWT is not running.
        /// This function performs the following operations:
        ///     1. Remove all WWT only layers.
        ///     2. Mark all Local IN WWT layer as not in sync with WWT.
        ///     3. Set the Selected layer map to null if the Selected WWT layer was deleted in WWT. 
        /// </summary>
        /// <param name="workbookMap">
        /// Parent container.
        /// </param>
        internal static void CleanUpWWTLayers(this WorkbookMap workbookMap)
        {
            if (workbookMap != null)
            {
                // 1. Create en empty collection, since this method is called only when WWT is not running. Empty collection is needed
                //    to be passed for UpdateGroupStatus method.
                ICollection<Group> groups = new List<Group>();
                LayerMap selectedLayerMap = workbookMap.SelectedLayerMap;

                // 2. Remove all WWT only layers.
                workbookMap.RemoveWWTLayers();

                // 3. On Cleanup, no WWT layers are connected with WWT.
                workbookMap.LocalLayerMaps.ForEach(
                    layer =>
                    {
                        // 3.a. Set IsNotInSync.
                        layer.IsNotInSync = true;

                        // 3.b. If the group (reference frame/ Layer Group) is deleted in WWT, then set the IsDeleted flag in Group to true.
                        UpdateGroupStatus(layer.LayerDetails.Group, groups);
                    });

                // 4. Set the Selected layer map to null if the Selected WWT layer was deleted in WWT. 
                //  If the selected layer is WWT layer/is localInWWT which is not deleted, this will be selected while 
                if (selectedLayerMap != null && selectedLayerMap.MapType == LayerMapType.WWT && !workbookMap.Exists(selectedLayerMap.LayerDetails.ID))
                {
                    workbookMap.SelectedLayerMap = null;
                }
            }
        }

        /// <summary>
        /// Removes the layer maps present in affected layers from the workbook map
        /// </summary>
        /// <param name="workbookMap">The workbook map which is in scope</param>
        /// <param name="affectedLayerList">The list of layers that need to be removed from the workbook map</param>
        internal static void RemoveAffectedLayers(this WorkbookMap workbookMap, List<LayerMap> affectedLayerList)
        {
            if (workbookMap != null && affectedLayerList != null)
            {
                foreach (LayerMap layerMap in affectedLayerList)
                {
                    if (workbookMap.Exists(layerMap.LayerDetails.ID))
                    {
                        workbookMap.AllLayerMaps.Remove(layerMap);
                    }
                }
            }
        }

        /// <summary>
        /// Get names of all the ranges for layers which are in sync
        /// </summary>
        /// <param name="workbookMap">current workbookMap</param>
        /// <returns>dictionary of range name and address</returns>
        internal static Dictionary<string, string> GetNamedRangesForInSyncLayers(this WorkbookMap workbookMap)
        {
            var allNamedRange = new Dictionary<string, string>();
            if (workbookMap != null)
            {
                // Get all Layer details into dictionary.
                workbookMap.LocalInWWTLayerMaps.ForEach(item =>
                {
                    // Only in sync layers
                    if (!item.IsNotInSync)
                    {
                        allNamedRange.Add(item.RangeDisplayName, item.RangeName.RefersTo as string);
                    }
                });
            }

            return allNamedRange;
        }

        /// <summary>
        /// Stops all notifications for all the layers of the current workbook map.
        /// </summary>
        /// <param name="workbookMap">WorkbookMap instance</param>
        internal static void StopAllNotifications(this WorkbookMap workbookMap)
        {
            if (workbookMap != null)
            {
                // Get all Layer details into dictionary.
                workbookMap.AllLayerMaps.ForEach(item =>
                {
                    item.IsNotInSync = true;
                });
            }
        }

        /// <summary>
        /// This function is used to search for the given group in the WWT Groups list.
        /// </summary>
        /// <param name="groupName">
        /// Name of the Group.
        /// </param>
        /// <param name="path">
        /// Path of the Group.
        /// </param>
        /// <param name="wwtGroups">
        /// List of WWT groups.
        /// </param>
        /// <returns>
        /// A Group that represents the group name and path.
        /// </returns>
        private static Group SearchGroup(string groupName, string path, ICollection<Group> wwtGroups)
        {
            Group result = null;
            foreach (Group group in wwtGroups)
            {
                if (string.CompareOrdinal(groupName, group.Name) == 0 &&
                                string.CompareOrdinal(path, group.Path) == 0)
                {
                    result = group;
                }
                else
                {
                    result = SearchGroup(groupName, path, group.Children);
                }

                if (result != null)
                {
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// This function is used to sync WWT layers into the workbook instance and 
        ///     then set the flag IsNotInSync for the layer deleted in WWT.
        /// </summary>
        /// <param name="workbookMap">
        /// Parent container.
        /// </param>
        /// <param name="groups">
        /// Collection of WWT groups.
        /// </param>
        private static void SyncWWTLayers(this WorkbookMap workbookMap, ICollection<Group> groups)
        {
            if (workbookMap != null)
            {
                // 1. Append all WWT layer which are not there in the LocalLayers.
                foreach (Group group in groups)
                {
                    AddWWTLayers(workbookMap, group);
                }

                // 2. Clean up all local layer in the current workbook map.
                workbookMap.LocalLayerMaps.ForEach(
                    layer =>
                    {
                        // 2.a. If the local layer is deleted in WWT, then set the IsNotInSync flag to true.
                        if (!WWTManager.IsValidLayer(layer.LayerDetails.ID, groups))
                        {
                            layer.IsNotInSync = true;
                        }

                        // 2.b. If the group (reference frame/ Layer Group) is deleted in WWT, then set the IsDeleted flag in Group to true.
                        UpdateGroupStatus(layer.LayerDetails.Group, groups);
                    });
            }
        }

        /// <summary>
        /// This function is used to update the status of the groups.
        /// </summary>
        /// <param name="group">
        /// Instance of the group which has to be updated.
        /// </param>
        /// <param name="wwtGroups">
        /// List of WWT groups.
        /// </param>
        private static void UpdateGroupStatus(Group group, ICollection<Group> wwtGroups)
        {
            group.IsDeleted = !WWTManager.IsValidGroup(group, wwtGroups);
            if (group.Parent != null)
            {
                UpdateGroupStatus(group.Parent, wwtGroups);
            }
        }

        /// <summary>
        /// This function is used to remove WWT layers.
        /// </summary>
        /// <param name="workbookMap">
        /// Parent container.
        /// </param>
        private static void RemoveWWTLayers(this WorkbookMap workbookMap)
        {
            if (workbookMap != null)
            {
                lock (workbookMap.AllLayerMaps)
                {
                    // Remove all WWT only layers.
                    workbookMap.WWTLayerMaps.ForEach(
                        wwtLayer =>
                        {
                            // This is needed to stop the notifications.
                            wwtLayer.IsNotInSync = true;
                            workbookMap.AllLayerMaps.Remove(wwtLayer);
                        });
                }
            }
        }

        /// <summary>
        /// This function is used to Add WWT layers.
        /// </summary>
        /// <param name="workbookMap">
        /// Parent container.
        /// </param>
        /// <param name="group">
        /// Group in which the layer is present.
        /// </param>
        private static void AddWWTLayers(WorkbookMap workbookMap, Group group)
        {
            foreach (string layerID in group.LayerIDs)
            {
                if (!workbookMap.Exists(layerID))
                {
                    Layer wwtLayer = WWTManager.GetLayerDetails(layerID, group, true);
                    if (wwtLayer != null)
                    {
                        workbookMap.AllLayerMaps.Add(new LayerMap(wwtLayer));
                    }
                }
            }

            foreach (Group child in group.Children)
            {
                AddWWTLayers(workbookMap, child);
            }
        }
    }
}