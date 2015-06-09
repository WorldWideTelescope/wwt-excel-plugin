//-----------------------------------------------------------------------
// <copyright file="LayerDetailsViewModelExtensions.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Extension methods for the view model
    /// </summary>
    internal static class LayerDetailsViewModelExtensions
    {
        /// <summary>
        /// Builds the layer details view model for the given workbook. Since rebuilding the reference frame dropdown makes call to 
        /// WWT API, in case if WWT is not running, caller will send false as it's value to avoid addition delay in showing results.
        /// </summary>
        /// <param name="layerDetailsModel">LayerDetailsViewModel object getting build</param>
        /// <param name="currentWorkbook">Current workbook object</param>
        /// <param name="rebuildReferenceFrameDropDown">Whether to rebuild reference frame dropdown or not?</param>
        /// <returns>Build LayerDetailsViewModel object</returns>
        internal static LayerDetailsViewModel BuildLayerDetailsViewModel(this LayerDetailsViewModel layerDetailsModel, WorkbookMap currentWorkbook, bool rebuildReferenceFrameDropDown)
        {
            // This is for building view model from workbook map. This will be used in Initialize, open & create new range scenarios
            if (layerDetailsModel != null)
            {
                // Build the view model for the drop down based on the layer map details
                layerDetailsModel.Layers = new ObservableCollection<LayerMapDropDownViewModel>();
                if (currentWorkbook != null)
                {
                    // Rebuilds the layer view model dropdown
                    WorkflowController.Instance.RebuildGroupLayerDropdown();

                    if (rebuildReferenceFrameDropDown)
                    {
                        WorkflowController.Instance.BuildReferenceFrameDropDown();
                    }

                    // If the current workbook has a selected layer, set current layer property which well set all the layer properties
                    if (currentWorkbook.SelectedLayerMap != null)
                    {
                        layerDetailsModel.Currentlayer = currentWorkbook.SelectedLayerMap;
                    }
                    else
                    {
                        LayerMapDropDownViewModel selectedLayerMap = layerDetailsModel.Layers.Where(layer => layer.Name.Equals(Properties.Resources.DefaultSelectedLayerName, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
                        if (selectedLayerMap != null)
                        {
                            layerDetailsModel.SelectedLayerText = selectedLayerMap.Name;
                        }
                        layerDetailsModel.Currentlayer = null;
                        layerDetailsModel.IsHelpTextVisible = true;
                        layerDetailsModel.IsTabVisible = false;
                    }
                }
                else
                {
                    layerDetailsModel.Layers.Add(new LayerMapDropDownViewModel() { Name = Properties.Resources.LocalLayerName });
                    layerDetailsModel.Layers.Add(new LayerMapDropDownViewModel() { Name = Properties.Resources.WWTLayerName });
                }
            }

            return layerDetailsModel;
        }

        internal static LayerDetailsViewModel SetLayerMap(this LayerDetailsViewModel layerDetailsModel, LayerMap selectedLayerMap)
        {
            // Set the ViewModel for Layer properties using the selected layer map from Core OM
            // This is one place from where all the View model properties are set and bound to task pane
            // This will be used when a layer is changed in the layer drop down
            if (selectedLayerMap != null)
            {
                if (selectedLayerMap.LayerDetails != null)
                {
                    LayerDetailsViewModel.IsPropertyChangedFromCode = true;

                    // Set all the user selections as well as default lists which are present only in view model
                    // Need to set properties which are in different format in View model from Model
                    layerDetailsModel.SelectedLayerName = selectedLayerMap.LayerDetails.Name;
                    layerDetailsModel.SelectedLayerText = LayerDetailsViewModel.GetLayerNameOnMapType(selectedLayerMap, selectedLayerMap.LayerDetails.Name);
                    layerDetailsModel = SetGroup(layerDetailsModel, selectedLayerMap);

                    // Binding the column data to the map columns
                    ObservableCollection<Column> columns = new ObservableCollection<Column>();
                    ColumnExtensions.PopulateColumnList().ToList().ForEach(col =>
                    {
                        columns.Add(col);
                    });

                    layerDetailsModel.ColumnsView = new ObservableCollection<ColumnViewModel>();
                    layerDetailsModel.SizeColumnList = new ObservableCollection<KeyValuePair<int, string>>();
                    layerDetailsModel.HoverTextColumnList = new ObservableCollection<KeyValuePair<int, string>>();

                    // Remove the columns based on the group selected.
                    layerDetailsModel.RemoveColumns(columns);

                    // Set the Mapped column type based on the group selected.
                    selectedLayerMap.UpdateMappedColumns();

                    // Validates if X,Y and Z columns are present and lat/long or RA/DEC columns are not present. If so,
                    // remove depth and alt from columns collection.
                    ValidateXYZ(selectedLayerMap, columns);

                    layerDetailsModel.PopulateColumns(selectedLayerMap, columns);

                    // Adding None to size column list
                    layerDetailsModel.SizeColumnList.Add(new KeyValuePair<int, string>(-1, Properties.Resources.NoneString));
                    layerDetailsModel.HoverTextColumnList.Add(new KeyValuePair<int, string>(-1, Properties.Resources.NoneString));

                    int index = 0;
                    foreach (string headerData in selectedLayerMap.HeaderRowData)
                    {
                        layerDetailsModel.SizeColumnList.Add(new KeyValuePair<int, string>(index, headerData));
                        layerDetailsModel.HoverTextColumnList.Add(new KeyValuePair<int, string>(index, headerData));
                        index++;
                    }

                    layerDetailsModel = SetSelectedSize(layerDetailsModel, selectedLayerMap);
                    layerDetailsModel = SetSelectedHoverText(layerDetailsModel, selectedLayerMap);

                    // Sets view in WWT visibility
                    layerDetailsModel.IsViewInWWTEnabled = (selectedLayerMap.MapType == LayerMapType.Local || (selectedLayerMap.MapType == LayerMapType.LocalInWWT && selectedLayerMap.IsNotInSync));
                    layerDetailsModel.IsLayerInSyncInfoVisible = (selectedLayerMap.MapType == LayerMapType.WWT || selectedLayerMap.MapType == LayerMapType.LocalInWWT) && !selectedLayerMap.IsNotInSync;

                    // On click of layer dropdown the callout need not be shown, the callout visibility is set to
                    // false on the click of the dropdown.
                    if (LayerDetailsViewModel.IsCallOutRequired)
                    {
                        layerDetailsModel.IsCallOutVisible = layerDetailsModel.IsViewInWWTEnabled;

                        if (layerDetailsModel.IsLayerInSyncInfoVisible)
                        {
                            // Start the animation for the layer in sync text.
                            layerDetailsModel.StartShowHighlightAnimationTimer();
                        }
                    }
                    else
                    {
                        layerDetailsModel.IsCallOutVisible = false;
                    }

                    // Sets if the custom task pane buttons are enabled
                    layerDetailsModel.IsShowRangeEnabled = GetButtonEnability(selectedLayerMap);
                    layerDetailsModel.IsDeleteMappingEnabled = GetButtonEnability(selectedLayerMap);
                    layerDetailsModel.IsGetLayerDataEnabled = (selectedLayerMap.MapType == LayerMapType.WWT || (selectedLayerMap.MapType == LayerMapType.LocalInWWT && !selectedLayerMap.IsNotInSync));

                    layerDetailsModel.IsUpdateLayerEnabled = GetButtonEnability(selectedLayerMap);
                    layerDetailsModel.IsReferenceGroupEnabled = selectedLayerMap.IsLayerCreated();
                    layerDetailsModel.IsFarSideShown = selectedLayerMap.LayerDetails.ShowFarSide;

                    SetDistanceUnit(layerDetailsModel, selectedLayerMap);

                    // Sets RAUnit visibility
                    layerDetailsModel.SetRAUnitVisibility();

                    layerDetailsModel.SetMarkerTabVisibility();

                    layerDetailsModel.SelectedFadeType = layerDetailsModel.FadeTypes.Where(fadetype => fadetype.Key == selectedLayerMap.LayerDetails.FadeType).FirstOrDefault();
                    layerDetailsModel.SelectedScaleRelative = layerDetailsModel.ScaleRelatives.Where(scaleRelative => scaleRelative.Key == selectedLayerMap.LayerDetails.MarkerScale).FirstOrDefault();
                    layerDetailsModel.SelectedScaleType = layerDetailsModel.ScaleTypes.Where(scaleType => scaleType.Key == selectedLayerMap.LayerDetails.PointScaleType).FirstOrDefault();
                    layerDetailsModel.SelectedMarkerType = layerDetailsModel.MarkerTypes.Where(markerType => markerType.Key == selectedLayerMap.LayerDetails.PlotType).FirstOrDefault();
                    layerDetailsModel.SelectedPushpinId = layerDetailsModel.PushPinTypes.Where(pushpin => pushpin.Key == selectedLayerMap.LayerDetails.MarkerIndex).FirstOrDefault();

                    layerDetailsModel.LayerOpacity.SelectedSliderValue = selectedLayerMap.LayerDetails.Opacity * 100;
                    layerDetailsModel.TimeDecay.SelectedSliderValue = GetSelectedTimeDecayValue(layerDetailsModel.TimeDecay, selectedLayerMap.LayerDetails.TimeDecay);
                    layerDetailsModel.ScaleFactor.SelectedSliderValue = GetSelectedTimeDecayValue(layerDetailsModel.ScaleFactor, selectedLayerMap.LayerDetails.ScaleFactor);

                    ////Set properties directly exposed from Model
                    layerDetailsModel.BeginDate = selectedLayerMap.LayerDetails.StartTime;
                    layerDetailsModel.EndDate = selectedLayerMap.LayerDetails.EndTime;
                    layerDetailsModel.FadeTime = selectedLayerMap.LayerDetails.FadeSpan.ToString();
                    layerDetailsModel.ColorBackground = LayerDetailsViewModel.ConvertToSolidColorBrush(selectedLayerMap.LayerDetails.Color);
                    LayerDetailsViewModel.IsPropertyChangedFromCode = false;
                }
            }

            return layerDetailsModel;
        }

        /// <summary>
        /// Get time decay value from layer details view model
        /// </summary>
        /// <param name="layerDetailsModel">layerDetailsModel instance</param>
        /// <returns>actual time decay</returns>
        internal static double GetActualTimeDecayValue(this LayerDetailsViewModel layerDetailsModel)
        {
            double timeDecay = 0;
            if (layerDetailsModel != null && layerDetailsModel.TimeDecay != null && layerDetailsModel.TimeDecay.SelectedSliderValue != 0)
            {
                timeDecay = layerDetailsModel.TimeDecay.SliderTicks[Convert.ToInt32(layerDetailsModel.TimeDecay.SelectedSliderValue) - 1];
            }

            return timeDecay;
        }

        /// <summary>
        /// Get scale factor value from layer details view model
        /// </summary>
        /// <param name="layerDetailsModel">layerDetailsModel instance</param>
        /// <returns>actual scale factor</returns>
        internal static double GetActualScaleFactorValue(this LayerDetailsViewModel layerDetailsModel)
        {
            double scaleFactor = 0;
            if (layerDetailsModel != null && layerDetailsModel.ScaleFactor != null && layerDetailsModel.ScaleFactor.SelectedSliderValue != 0)
            {
                scaleFactor = layerDetailsModel.ScaleFactor.SliderTicks[Convert.ToInt32(layerDetailsModel.ScaleFactor.SelectedSliderValue) - 1];
            }

            return scaleFactor;
        }

        #region Private methods

        /// <summary>
        /// Gets the selected time decay value
        /// </summary>
        /// <param name="sliderViewModel">Slider view model</param>
        /// <param name="selectedTimeDecay">selected time decay</param>
        /// <returns>Selected time decay value</returns>
        private static double GetSelectedTimeDecayValue(SliderViewModel sliderViewModel, double selectedTimeDecay)
        {
            int index = sliderViewModel.SliderTicks.IndexOf(selectedTimeDecay);
            return index + 1;
        }

        /// <summary>
        /// Gets if the button is enable on the basis of the layer being local/local in WWT
        /// </summary>
        /// <param name="selectedLayerMap">Selected layer map</param>
        /// <returns>If the button has to be enabled or not</returns>
        private static bool GetButtonEnability(LayerMap selectedLayerMap)
        {
            return (selectedLayerMap.MapType == LayerMapType.Local || selectedLayerMap.MapType == LayerMapType.LocalInWWT);
        }

        /// <summary>
        /// Sets group for the view model
        /// </summary>
        /// <param name="layerDetailsModel">Layer details view model</param>
        /// <param name="selectedLayerMap">Selected layer map</param>
        /// <returns>Updated layer view model</returns>
        private static LayerDetailsViewModel SetGroup(LayerDetailsViewModel layerDetailsModel, LayerMap selectedLayerMap)
        {
            if (selectedLayerMap.LayerDetails.Group != null)
            {
                layerDetailsModel.SelectedGroupText = selectedLayerMap.LayerDetails.Group.Name;
                layerDetailsModel.SelectedGroup = selectedLayerMap.LayerDetails.Group;
            }
            return layerDetailsModel;
        }

        /// <summary>
        /// Sets distance unit visibility and selected distance unit.
        /// </summary>
        /// <param name="layerDetailsModel">Layer details view model</param>
        /// <param name="selectedLayerMap">Selected layer map</param>
        /// <returns>Updated layer view model</returns>
        private static LayerDetailsViewModel SetDistanceUnit(LayerDetailsViewModel layerDetailsModel, LayerMap selectedLayerMap)
        {
            if (layerDetailsModel.ColumnsView != null && layerDetailsModel.ColumnsView.Count > 0)
            {
                ColumnViewModel depthColumn = layerDetailsModel.ColumnsView.Where(columnValue => columnValue.SelectedWWTColumn.IsDepthColumn()).FirstOrDefault();
                ColumnViewModel xyzColumn = layerDetailsModel.ColumnsView.Where(columnValue => columnValue.SelectedWWTColumn.IsXYZColumn()).FirstOrDefault();
                if (depthColumn != null || xyzColumn != null)
                {
                    layerDetailsModel.IsDistanceVisible = true;
                }
                else
                {
                    layerDetailsModel.IsDistanceVisible = false;
                }
                if (layerDetailsModel.DistanceUnits != null && layerDetailsModel.DistanceUnits.Count > 0)
                {
                    layerDetailsModel.SelectedDistanceUnit = layerDetailsModel.DistanceUnits.Where(distanceUnit => distanceUnit.Key == selectedLayerMap.LayerDetails.AltUnit).FirstOrDefault();
                }
            }
            return layerDetailsModel;
        }

        /// <summary>
        /// Sets selected size based on the layer properties
        /// </summary>
        /// <param name="layerDetailsModel">Layer details view model</param>
        /// <param name="selectedLayerMap">Selected layer map</param>
        /// <returns>Updated layer view model</returns>
        private static LayerDetailsViewModel SetSelectedSize(LayerDetailsViewModel layerDetailsModel, LayerMap selectedLayerMap)
        {
            if (selectedLayerMap.LayerDetails.SizeColumn != Constants.DefaultColumnIndex)
            {
                layerDetailsModel.SelectedSize = layerDetailsModel.SizeColumnList.Where(item => item.Key == selectedLayerMap.LayerDetails.SizeColumn).FirstOrDefault();
            }
            else
            {
                layerDetailsModel.SelectedSize = layerDetailsModel.SizeColumnList[0];
            }
            return layerDetailsModel;
        }

        /// <summary>
        ///  Sets selected hover text based on the layer properties
        /// </summary>
        /// <param name="layerDetailsModel">Layer details view model</param>
        /// <param name="selectedLayerMap">Selected layer map</param>
        /// <returns>Updated layer view model</returns>
        private static LayerDetailsViewModel SetSelectedHoverText(LayerDetailsViewModel layerDetailsModel, LayerMap selectedLayerMap)
        {
            if (selectedLayerMap.LayerDetails.NameColumn != Constants.DefaultColumnIndex)
            {
                layerDetailsModel.SelectedHoverText = layerDetailsModel.HoverTextColumnList.Where(item => item.Key == selectedLayerMap.LayerDetails.NameColumn).FirstOrDefault();
            }
            else
            {
                layerDetailsModel.SelectedHoverText = layerDetailsModel.HoverTextColumnList[0];
            }
            return layerDetailsModel;
        }

        /// <summary>
        /// Validates if X,Y and z column is present and lat/long or RA/DEC column is not present if so 
        /// remove depth and alt from columns collection
        /// </summary>
        private static void ValidateXYZ(LayerMap selectedLayerMap, Collection<Column> columns)
        {
            if (selectedLayerMap.IsXYZLayer())
            {
                columns.Remove(columns.Where(col => col.ColType == ColumnType.Depth).FirstOrDefault());
                columns.Remove(columns.Where(col => col.ColType == ColumnType.Alt).FirstOrDefault());
                columns.Remove(columns.Where(col => col.ColType == ColumnType.Distance).FirstOrDefault());
            }
        }

        #endregion
    }
}
