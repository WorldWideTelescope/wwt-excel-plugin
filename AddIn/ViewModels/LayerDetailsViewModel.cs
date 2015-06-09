//-----------------------------------------------------------------------
// <copyright file="LayerDetailsViewModel.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Research.Wwt.Excel.Addin.Properties;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// View model for layer details shown in the layer manager
    /// </summary>
    public class LayerDetailsViewModel : PropertyChangeBase
    {
        #region Private Properties

        private LayerMap currentLayer;

        private bool isTabVisible;
        private bool isHelpTextVisible;
        private bool isViewInWWTEnabled;
        private bool isMarkerTabEnabled;
        private bool isDistanceVisible;
        private bool isCallOutVisible = true;
        private bool isShowRangeEnabled;
        private bool isDeleteMappingEnabled;
        private bool isGetLayerDataEnabled;
        private bool isUpdateLayerEnabled;
        private bool isReferenceGroupEnabled;
        private bool isRAUnitVisible;
        private bool isFarSideShown;
        private bool isRenderingTimeoutAlertShown;
        private bool isLayerInSyncInfoVisible;
        private bool pushpinMarkerTypeSelected;

        private ObservableCollection<KeyValuePair<int, string>> sizeColumnList;
        private ObservableCollection<KeyValuePair<int, string>> hoverTextColumnList;
        private ObservableCollection<ColumnViewModel> column;
        private ObservableCollection<GroupViewModel> referenceGroups;
        private ObservableCollection<LayerMapDropDownViewModel> layers;

        private Collection<KeyValuePair<MarkerType, string>> markerTypes;
        private Collection<KeyValuePair<int, BitmapImage>> pushpinTypes;
        private Collection<KeyValuePair<ScaleRelativeType, string>> scaleRelatives;
        private Collection<KeyValuePair<ScaleType, string>> scaleTypes;
        private Collection<KeyValuePair<FadeType, string>> fadetypes;
        private Collection<KeyValuePair<AltUnit, string>> distanceUnits;
        private Collection<KeyValuePair<AngleUnit, string>> rightAscentionUnits;

        private KeyValuePair<FadeType, string> selectedFadeType;
        private KeyValuePair<ScaleType, string> selectedScaleType;
        private KeyValuePair<ScaleRelativeType, string> selectedScaleRelative;
        private KeyValuePair<AltUnit, string> selectedDistanceUnit;
        private KeyValuePair<AngleUnit, string> selectedRAUnit;
        private KeyValuePair<MarkerType, string> selectedMarkerType;
        private KeyValuePair<int, BitmapImage> selectedPushpinId;

        private int selectedTabIndex;
        private string selectedLayerName;
        private string layerDataDisplayName;
        private SliderViewModel timeDecay;
        private SliderViewModel layerOpacity;
        private SliderViewModel scaleFactor;
        private string selectedGroupText;
        private Group selectedGroup;
        private string selectedLayerText;

        private KeyValuePair<int, string> selectedSizeColumn;
        private KeyValuePair<int, string> selectedHoverText;

        private LayerMapDropDownViewModel selectedLayerMapDropDown;
        private DownloadUpdatesViewModel downloadUpdatesViewModelInstance;

        private ICommand selectionCommand;
        private ICommand controlCommand;
        private ICommand viewInWWTCommand;
        private ICommand colorPalletCommand;
        private ICommand callOutCommand;
        private ICommand showRangeCommand;
        private ICommand deleteMappingCommand;
        private ICommand getLayerDataCommand;
        private ICommand updateLayerCommand;
        private ICommand layerMapNameChangeCommand;
        private ICommand fadeTimeChangeCommand;
        private ICommand refreshDropDownCommand;
        private ICommand refreshGroupDropDownCommand;
        private ICommand sizeColumnChangeCommand;
        private ICommand downloadUpdatesCommand;
        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the LayerDetailsViewModel class
        /// </summary>
        public LayerDetailsViewModel()
        {
            this.AttachCommands();
            this.SetDefaultValues();
            this.ReferenceGroups = new ObservableCollection<GroupViewModel>();
        }

        #endregion

        #region CustomEvent

        /// <summary>
        /// Event is fired on any state change on the custom task pane.
        /// </summary>
        public event EventHandler CustomTaskPaneStateChangedEvent;

        /// <summary>
        /// Event is fired on any state change on the custom task pane.
        /// </summary>
        public event EventHandler LayerSelectionChangedEvent;

        /// <summary>
        /// Event is fired on any state change on the custom task pane.
        /// </summary>
        public event EventHandler ViewnInWWTClickedEvent;

        /// <summary>
        /// Event is fired when we need to select the range for the selected layer.
        /// </summary>
        public event EventHandler ShowRangeClickedEvent;

        /// <summary>
        /// Event is fired when we need to delete mappings for the selected layer.
        /// </summary>
        public event EventHandler DeleteMappingClickedEvent;

        /// <summary>
        /// Event is fired when we need to get layer data for WWT or local in WWT layer
        /// </summary>
        public event EventHandler GetLayerDataClickedEvent;

        /// <summary>
        /// Event is fired when we need to update the range of selected layer.
        /// </summary>
        public event EventHandler UpdateLayerClickedEvent;

        /// <summary>
        /// Event is fired when the dropdown for refresh is clicked
        /// </summary>
        public event EventHandler RefreshDropDownClickedEvent;

        /// <summary>
        /// Event is fired when the group dropdown is refreshed.
        /// </summary>
        public event EventHandler RefreshGroupDropDownClickedEvent;

        /// <summary>
        /// Event is fired when the reference group selection changes from Sky to planet 
        /// or vice-versa
        /// </summary>
        public event EventHandler ReferenceSelectionChanged;

        /// <summary>
        /// Event is fired when the download updates button is clicked
        /// </summary>
        public event EventHandler DownloadUpdatesClickedEvent;
        #endregion

        #region Properties

        #region Boolean Properties

        /// <summary>
        /// Gets or sets a value indicating whether the property change is happening from code or not
        /// </summary>
        public static bool IsPropertyChangedFromCode { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the call out is to be shown or not.
        /// If the layer dropdown is clicked, callout is not required.
        /// </summary>
        public static bool IsCallOutRequired { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether Tab is Visible or not
        /// </summary>
        public bool IsTabVisible
        {
            get
            {
                return this.isTabVisible;
            }
            set
            {
                isTabVisible = value;
                OnPropertyChanged("IsTabVisible");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether Help text is Visible or not
        /// </summary>
        public bool IsHelpTextVisible
        {
            get
            {
                return this.isHelpTextVisible;
            }
            set
            {
                isHelpTextVisible = value;
                OnPropertyChanged("IsHelpTextVisible");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether distance property is visible or not
        /// </summary>
        public bool IsDistanceVisible
        {
            get
            {
                return this.isDistanceVisible;
            }
            set
            {
                isDistanceVisible = value;
                OnPropertyChanged("IsDistanceVisible");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether view in WWT button is enabled or not
        /// </summary>
        public bool IsViewInWWTEnabled
        {
            get
            {
                return this.isViewInWWTEnabled;
            }
            set
            {
                isViewInWWTEnabled = value;
                OnPropertyChanged("IsViewInWWTEnabled");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether marker tab is enabled or not
        /// </summary>
        public bool IsMarkerTabEnabled
        {
            get
            {
                return this.isMarkerTabEnabled;
            }
            set
            {
                isMarkerTabEnabled = value;
                OnPropertyChanged("IsMarkerTabEnabled");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether callout is visible or not
        /// </summary>
        public bool IsCallOutVisible
        {
            get
            {
                return this.isCallOutVisible;
            }
            set
            {
                this.isCallOutVisible = value;

                // If the callout visibility is set to true start the timer
                if (this.isCallOutVisible)
                {
                    this.StartCallOutVisibilityTimer();
                }
                OnPropertyChanged("IsCallOutVisible");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether show range is enabled or not
        /// </summary>
        public bool IsShowRangeEnabled
        {
            get
            {
                return this.isShowRangeEnabled;
            }
            set
            {
                this.isShowRangeEnabled = value;
                OnPropertyChanged("IsShowRangeEnabled");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether delete mapping is enabled or not
        /// </summary>
        public bool IsDeleteMappingEnabled
        {
            get
            {
                return this.isDeleteMappingEnabled;
            }
            set
            {
                this.isDeleteMappingEnabled = value;
                OnPropertyChanged("IsDeleteMappingEnabled");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether get layer data is enabled or not
        /// </summary>
        public bool IsGetLayerDataEnabled
        {
            get
            {
                return this.isGetLayerDataEnabled;
            }
            set
            {
                this.isGetLayerDataEnabled = value;
                OnPropertyChanged("IsGetLayerDataEnabled");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether update layer data is enabled or not
        /// </summary>
        public bool IsUpdateLayerEnabled
        {
            get
            {
                return this.isUpdateLayerEnabled;
            }
            set
            {
                this.isUpdateLayerEnabled = value;
                OnPropertyChanged("IsUpdateLayerEnabled");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether RA Unit is visible or not
        /// </summary>
        public bool IsRAUnitVisible
        {
            get
            {
                return this.isRAUnitVisible;
            }
            set
            {
                this.isRAUnitVisible = value;
                OnPropertyChanged("IsRAUnitVisible");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether far side is shown or not
        /// </summary>
        public bool IsFarSideShown
        {
            get
            {
                return this.isFarSideShown;
            }
            set
            {
                this.isFarSideShown = value;
                OnPropertyChanged("IsFarSideShown");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether RenderingTimeout alert is shown or not
        /// </summary>
        public bool IsRenderingTimeoutAlertShown
        {
            get
            {
                return this.isRenderingTimeoutAlertShown;
            }
            set
            {
                this.isRenderingTimeoutAlertShown = value;
                OnPropertyChanged("IsRenderingTimeoutAlertShown");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether layer is in synch with WWT
        /// </summary>
        public bool IsLayerInSyncInfoVisible
        {
            get
            {
                return this.isLayerInSyncInfoVisible;
            }
            set
            {
                this.isLayerInSyncInfoVisible = value;
                OnPropertyChanged("IsLayerInSyncInfoVisible");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether Pushpin marker type is selected or not.
        /// </summary>
        public bool PushpinMarkerTypeSelected
        {
            get
            {
                return this.pushpinMarkerTypeSelected;
            }
            set
            {
                this.pushpinMarkerTypeSelected = value;
                OnPropertyChanged("PushpinMarkerTypeSelected");
            }
        }
        #endregion

        #region Collection Properties
        /// <summary>
        /// Gets or sets the instance of the DownloadUpdatesViewModel
        /// </summary>
        public DownloadUpdatesViewModel DownloadUpdatesViewModelInstance
        {
            get
            {
                return this.downloadUpdatesViewModelInstance;
            }
            set
            {
                this.downloadUpdatesViewModelInstance = value;
                OnPropertyChanged("DownloadUpdatesViewModelInstance");
            }
        }

        /// <summary>
        /// Gets Fade types
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures", Justification = "This will save creation of additional view model class with Id and Display value")]
        public ReadOnlyCollection<KeyValuePair<FadeType, string>> FadeTypes
        {
            get
            {
                return new ReadOnlyCollection<KeyValuePair<FadeType, string>>(this.fadetypes);
            }
        }

        /// <summary>
        /// Gets or sets Layers
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Observable collection for binding in XAML")]
        public ObservableCollection<LayerMapDropDownViewModel> Layers
        {
            get
            {
                return this.layers;
            }
            set
            {
                if (value != null)
                {
                    this.layers = value;
                    this.layers.CollectionChanged += new NotifyCollectionChangedEventHandler(LayersCollectionChanged);
                }
            }
        }

        /// <summary>
        /// Gets the marker type list
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures", Justification = "This will save creation of additional view model class with Id and Display value")]
        public ReadOnlyCollection<KeyValuePair<MarkerType, string>> MarkerTypes
        {
            get
            {
                return new ReadOnlyCollection<KeyValuePair<MarkerType, string>>(this.markerTypes);
            }
        }

        /// <summary>
        /// Gets the PushPin list
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures", Justification = "This will save creation of additional view model class with Id and Display value")]
        public ReadOnlyCollection<KeyValuePair<int, BitmapImage>> PushPinTypes
        {
            get
            {
                return new ReadOnlyCollection<KeyValuePair<int, BitmapImage>>(this.pushpinTypes);
            }
        }

        /// <summary>
        /// Gets Scale types
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures", Justification = "This will save creation of additional view model class with Id and Display value")]
        public ReadOnlyCollection<KeyValuePair<ScaleType, string>> ScaleTypes
        {
            get
            {
                return new ReadOnlyCollection<KeyValuePair<ScaleType, string>>(this.scaleTypes);
            }
        }

        /// <summary>
        /// Gets or sets SizeColumnList
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Observable collection for binding in XAML")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures", Justification = "This will save creation of additional view model class with Id and Display value")]
        public ObservableCollection<KeyValuePair<int, string>> SizeColumnList
        {
            get
            {
                return this.sizeColumnList;
            }

            set
            {
                this.sizeColumnList = value;
                OnPropertyChanged("SizeColumnList");
            }
        }

        /// <summary>
        /// Gets or sets SizeColumnList
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Observable collection for binding in XAML")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures", Justification = "This will save creation of additional view model class with Id and Display value")]
        public ObservableCollection<KeyValuePair<int, string>> HoverTextColumnList
        {
            get
            {
                return this.hoverTextColumnList;
            }

            set
            {
                this.hoverTextColumnList = value;
                OnPropertyChanged("HoverTextColumnList");
            }
        }

        /// <summary>
        /// Gets ScaleRelatives
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures", Justification = "This will save creation of additional view model class with Id and Display value")]
        public ReadOnlyCollection<KeyValuePair<ScaleRelativeType, string>> ScaleRelatives
        {
            get
            {
                return new ReadOnlyCollection<KeyValuePair<ScaleRelativeType, string>>(this.scaleRelatives);
            }
        }

        /// <summary>
        /// Gets or sets ColumnsView
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Observable collection for binding in XAML")]
        public ObservableCollection<ColumnViewModel> ColumnsView
        {
            get
            {
                return this.column;
            }

            set
            {
                if (value != null)
                {
                    this.column = value;
                    this.column.CollectionChanged += new System.Collections.Specialized.NotifyCollectionChangedEventHandler(MapColumnCollectionChanged);
                }
                OnPropertyChanged("ColumnsView");
            }
        }

        /// <summary>
        /// Gets DistanceUnits
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures", Justification = "This will save creation of additional view model class with Id and Display value")]
        public ReadOnlyCollection<KeyValuePair<AltUnit, string>> DistanceUnits
        {
            get
            {
                return new ReadOnlyCollection<KeyValuePair<AltUnit, string>>(this.distanceUnits);
            }
        }

        /// <summary>
        /// Gets or sets the layer groups/ reference frame
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Observable collection for binding in XAML")]
        public ObservableCollection<GroupViewModel> ReferenceGroups
        {
            get
            {
                return this.referenceGroups;
            }
            set
            {
                if (value != null)
                {
                    this.referenceGroups = value;
                    this.referenceGroups.CollectionChanged += new NotifyCollectionChangedEventHandler(GroupCollectionChanged);
                }
            }
        }

        /// <summary>
        /// Gets RA units
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures", Justification = "This will save creation of additional view model class with Id and Display value")]
        public ReadOnlyCollection<KeyValuePair<AngleUnit, string>> RightAscentionUnits
        {
            get
            {
                return new ReadOnlyCollection<KeyValuePair<AngleUnit, string>>(this.rightAscentionUnits);
            }
        }
        #endregion

        #region Layer Properties
        /// <summary>
        /// Gets or sets BeginDate
        /// </summary>
        public DateTime BeginDate
        {
            get
            {
                if (currentLayer != null && currentLayer.LayerDetails != null)
                {
                    return this.currentLayer.LayerDetails.StartTime;
                }
                return DateTime.MinValue;
            }
            set
            {
                if (currentLayer != null && currentLayer.LayerDetails != null)
                {
                    if (ValidateDate(value, this.EndDate))
                    {
                        currentLayer.LayerDetails.StartTime = value;
                        this.OnCustomTaskPaneStateChanged();
                    }
                }
                OnPropertyChanged("BeginDate");
            }
        }

        /// <summary>
        /// Gets or sets EndDate
        /// </summary>
        public DateTime EndDate
        {
            get
            {
                if (currentLayer != null && currentLayer.LayerDetails != null)
                {
                    return this.currentLayer.LayerDetails.EndTime;
                }
                return DateTime.MaxValue;
            }
            set
            {
                if (currentLayer != null && currentLayer.LayerDetails != null)
                {
                    if (ValidateDate(this.BeginDate, value))
                    {
                        this.currentLayer.LayerDetails.EndTime = value;
                        this.OnCustomTaskPaneStateChanged();
                    }
                }
                OnPropertyChanged("EndDate");
            }
        }

        /// <summary>
        /// Gets or sets FadeTime
        /// </summary>
        public string FadeTime
        {
            get
            {
                string fadespan = string.Empty;
                if (currentLayer != null && currentLayer.LayerDetails != null)
                {
                    fadespan = this.currentLayer.LayerDetails.FadeSpan.ToString();
                }
                return fadespan;
            }
            set
            {
                if (currentLayer != null && currentLayer.LayerDetails != null)
                {
                    TimeSpan timeSpan = new TimeSpan();
                    if (TimeSpan.TryParse(value, out timeSpan))
                    {
                        this.currentLayer.LayerDetails.FadeSpan = timeSpan;
                        this.OnCustomTaskPaneStateChanged();
                        OnPropertyChanged("FadeTime");
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets ColorBackground
        /// </summary>
        public Brush ColorBackground
        {
            get
            {
                if (currentLayer != null && currentLayer.LayerDetails != null)
                {
                    return ConvertToSolidColorBrush(this.currentLayer.LayerDetails.Color);
                }
                return null;
            }
            set
            {
                if (currentLayer != null && currentLayer.LayerDetails != null)
                {
                    this.currentLayer.LayerDetails.Color = ConvertColorToString(value);
                    this.OnCustomTaskPaneStateChanged();
                }
                OnPropertyChanged("ColorBackground");
            }
        }

        #endregion

        #region ICommand

        /// <summary>
        /// Gets Selection command
        /// </summary>
        public ICommand SelectionCommand
        {
            get { return selectionCommand; }
        }

        /// <summary>
        /// Gets map column command
        /// </summary>
        public ICommand ControlCommand
        {
            get { return this.controlCommand; }
        }

        /// <summary>
        /// Gets view in WWT command
        /// </summary>
        public ICommand ViewInWWTCommand
        {
            get { return this.viewInWWTCommand; }
        }

        /// <summary>
        /// Gets the color pallet command
        /// </summary>
        public ICommand ColorPalletCommand
        {
            get { return this.colorPalletCommand; }
        }

        /// <summary>
        /// Gets the callout command
        /// </summary>
        public ICommand CallOutCommand
        {
            get { return this.callOutCommand; }
        }

        /// <summary>
        /// Gets the show range command.
        /// </summary>
        public ICommand ShowRangeCommand
        {
            get { return this.showRangeCommand; }
        }

        /// <summary>
        /// Gets the delete mapping command.
        /// </summary>
        public ICommand DeleteMappingCommand
        {
            get { return this.deleteMappingCommand; }
        }

        /// <summary>
        /// Gets the layer data command
        /// </summary>
        public ICommand GetLayerDataCommand
        {
            get { return this.getLayerDataCommand; }
        }

        /// <summary>
        /// Gets the Update Layer command.
        /// </summary>
        public ICommand UpdateLayerCommand
        {
            get { return this.updateLayerCommand; }
        }

        /// <summary>
        /// Gets the layer name change command.
        /// </summary>
        public ICommand LayerMapNameChangeCommand
        {
            get { return this.layerMapNameChangeCommand; }
        }

        /// <summary>
        /// Gets the fade time change command.
        /// </summary>
        public ICommand FadeTimeChangeCommand
        {
            get { return this.fadeTimeChangeCommand; }
        }

        /// <summary>
        /// Gets the refresh command.
        /// </summary>
        public ICommand RefreshDropDownCommand
        {
            get { return this.refreshDropDownCommand; }
        }

        /// <summary>
        /// Gets the refresh group dropdown command.
        /// </summary>
        public ICommand RefreshGroupDropDownCommand
        {
            get { return this.refreshGroupDropDownCommand; }
        }

        /// <summary>
        /// Gets the size column changed command
        /// </summary>
        public ICommand SizeColumnChangeCommand
        {
            get { return this.sizeColumnChangeCommand; }
        }

        /// <summary>
        /// Gets the download updates command
        /// </summary>
        public ICommand DownloadUpdatesCommand
        {
            get { return this.downloadUpdatesCommand; }
        }

        #endregion

        #region Selected Properties

        /// <summary>
        /// Gets or sets the index of the selected tab in the layer manager pane
        /// </summary>
        public int SelectedTabIndex
        {
            get
            {
                return this.selectedTabIndex;
            }
            set
            {
                this.selectedTabIndex = value;
                OnPropertyChanged("SelectedTabIndex");
            }
        }

        /// <summary>
        /// Gets or sets SelectedScaleRelative
        /// </summary>
        public KeyValuePair<ScaleRelativeType, string> SelectedScaleRelative
        {
            get
            {
                return this.selectedScaleRelative;
            }
            set
            {
                selectedScaleRelative = value;
                OnPropertyChanged("SelectedScaleRelative");
            }
        }

        /// <summary>
        /// Gets or sets SelectedFadeType
        /// </summary>
        public KeyValuePair<FadeType, string> SelectedFadeType
        {
            get
            {
                return this.selectedFadeType;
            }
            set
            {
                selectedFadeType = value;
                OnPropertyChanged("SelectedFadeType");
            }
        }

        /// <summary>
        /// Gets or sets SelectedDistanceUnit
        /// </summary>
        public KeyValuePair<AltUnit, string> SelectedDistanceUnit
        {
            get
            {
                return this.selectedDistanceUnit;
            }
            set
            {
                selectedDistanceUnit = value;
                OnPropertyChanged("SelectedDistanceUnit");
            }
        }

        /// <summary>
        /// Gets or sets SelectedScaleType
        /// </summary>
        public KeyValuePair<ScaleType, string> SelectedScaleType
        {
            get
            {
                return this.selectedScaleType;
            }
            set
            {
                selectedScaleType = value;
                OnPropertyChanged("SelectedScaleType");
            }
        }

        /// <summary>
        /// Gets or sets SelectedLayerName
        /// </summary>
        public string SelectedLayerName
        {
            get
            {
                return this.selectedLayerName;
            }
            set
            {
                if (!string.IsNullOrWhiteSpace(value))
                {
                    selectedLayerName = value.Trim();
                    SetLayerName(value.Trim());
                }
                OnPropertyChanged("SelectedLayerName");
            }
        }

        /// <summary>
        /// Gets TimeDecay
        /// </summary>
        public SliderViewModel TimeDecay
        {
            get
            {
                return this.timeDecay;
            }
        }

        /// <summary>
        /// Gets LayerOpacity
        /// </summary>
        public SliderViewModel LayerOpacity
        {
            get
            {
                return this.layerOpacity;
            }
        }

        /// <summary>
        /// Gets ScaleFactor
        /// </summary>
        public SliderViewModel ScaleFactor
        {
            get
            {
                return this.scaleFactor;
            }
        }

        /// <summary>
        /// Gets or sets SelectedSize
        /// </summary>
        public KeyValuePair<int, string> SelectedSize
        {
            get
            {
                return this.selectedSizeColumn;
            }
            set
            {
                this.selectedSizeColumn = value;
                OnPropertyChanged("SelectedSize");
            }
        }

        /// <summary>
        /// Gets or sets SelectedHoverText
        /// </summary>
        public KeyValuePair<int, string> SelectedHoverText
        {
            get
            {
                return this.selectedHoverText;
            }
            set
            {
                this.selectedHoverText = value;
                OnPropertyChanged("SelectedHoverText");
            }
        }

        /// <summary>
        /// Gets or sets the selected layer in the dropdown
        /// </summary>
        public LayerMapDropDownViewModel SelectedLayerMapDropDown
        {
            get
            {
                return selectedLayerMapDropDown;
            }
            set
            {
                selectedLayerMapDropDown = value;
                OnPropertyChanged("SelectedLayerMapDropDown");
            }
        }

        /// <summary>
        /// Gets or sets layer data display name
        /// </summary>
        public string LayerDataDisplayName
        {
            get
            {
                return this.layerDataDisplayName;
            }
            set
            {
                this.layerDataDisplayName = value;
                OnPropertyChanged("LayerDataDisplayName");
            }
        }

        /// <summary>
        /// Gets or sets the layer group's text value 
        /// </summary>
        public string SelectedGroupText
        {
            get
            {
                return this.selectedGroupText;
            }
            set
            {
                this.selectedGroupText = value;
                OnPropertyChanged("SelectedGroupText");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the reference group should be enabled
        /// </summary>
        public bool IsReferenceGroupEnabled
        {
            get
            {
                return this.isReferenceGroupEnabled;
            }
            set
            {
                this.isReferenceGroupEnabled = value;
                OnPropertyChanged("IsReferenceGroupEnabled");
            }
        }

        /// <summary>
        /// Gets or sets the selected group
        /// </summary>
        public Group SelectedGroup
        {
            get
            {
                return this.selectedGroup;
            }
            set
            {
                this.selectedGroup = value;
                OnPropertyChanged("SelectedGroup");
            }
        }

        /// <summary>
        /// Gets or sets selected layer text
        /// </summary>
        public string SelectedLayerText
        {
            get
            {
                return this.selectedLayerText;
            }
            set
            {
                this.selectedLayerText = value;
                OnPropertyChanged("SelectedLayerText");
            }
        }

        /// <summary>
        /// Gets or sets selected RA unit
        /// </summary>
        public KeyValuePair<AngleUnit, string> SelectedRAUnit
        {
            get
            {
                return this.selectedRAUnit;
            }
            set
            {
                selectedRAUnit = value;
                OnPropertyChanged("SelectedRAUnit");
            }
        }

        /// <summary>
        /// Gets or sets the SelectedMarkerType
        /// </summary>
        public KeyValuePair<MarkerType, string> SelectedMarkerType
        {
            get
            {
                return this.selectedMarkerType;
            }
            set
            {
                if (value.Key == MarkerType.PushPin)
                {
                    this.PushpinMarkerTypeSelected = true;
                }
                else
                {
                    this.PushpinMarkerTypeSelected = false;
                }

                this.selectedMarkerType = value;
                OnPropertyChanged("SelectedMarkerType");
            }
        }

        /// <summary>
        /// Gets or sets the SelectedPushpinId
        /// </summary>
        public KeyValuePair<int, BitmapImage> SelectedPushpinId
        {
            get
            {
                return this.selectedPushpinId;
            }
            set
            {
                this.selectedPushpinId = value;
                OnPropertyChanged("SelectedPushpinId");
            }
        }

        /// <summary>
        /// Gets or sets current layer
        /// </summary>
        internal LayerMap Currentlayer
        {
            get
            {
                return currentLayer;
            }
            set
            {
                currentLayer = value;
                BindDatatoViewModel();
            }
        }

        #endregion

        #endregion

        #region Public and internal methods

        /// <summary>
        /// Converts the byte array to solid color brush
        /// </summary>
        /// <param name="colorArgb">ARGB color(byte array)</param>
        /// <returns>Solid color brush</returns>
        public static SolidColorBrush ConvertToSolidColorBrush(string colorArgb)
        {
            SolidColorBrush color = new SolidColorBrush(System.Windows.Media.Color.FromArgb(System.Drawing.Color.Red.A, System.Drawing.Color.Red.R, System.Drawing.Color.Red.G, System.Drawing.Color.Red.B));
            if (!string.IsNullOrEmpty(colorArgb))
            {
                colorArgb = colorArgb.Substring(colorArgb.IndexOf(":", StringComparison.OrdinalIgnoreCase) + 1);
                string[] colors = colorArgb.Split(':');
                color = new SolidColorBrush(Color.FromArgb(Convert.ToByte(colors[0], CultureInfo.InvariantCulture), Convert.ToByte(colors[1], CultureInfo.InvariantCulture), Convert.ToByte(colors[2], CultureInfo.InvariantCulture), Convert.ToByte(colors[3], CultureInfo.InvariantCulture)));
            }

            return color;
        }

        /// <summary>
        /// Raises event on custom task pane changed
        /// </summary>
        public void OnCustomTaskPaneStateChanged()
        {
            if (!LayerDetailsViewModel.IsPropertyChangedFromCode)
            {
                this.CustomTaskPaneStateChangedEvent.OnFire(this, new EventArgs());
            }
        }

        /// <summary>
        /// Set column collection based on selected reference frame
        /// </summary>
        /// <param name="columns">Collection of columns</param>
        public void RemoveColumns(Collection<Column> columns)
        {
            List<Column> removeColumns = null;
            if (this.SelectedGroup != null)
            {
                if (this.SelectedGroup.IsPlanet())
                {
                    removeColumns = columns.Where(columnVal => columnVal.ColType == ColumnType.RA || columnVal.ColType == ColumnType.Dec).ToList();
                }
                else
                {
                    removeColumns = columns.Where(columnVal => columnVal.ColType == ColumnType.Lat || columnVal.ColType == ColumnType.Long).ToList();
                }
                removeColumns.ForEach(column => columns.Remove(column));
            }
        }

        /// <summary>
        /// Sets distance unit visibility according to the selected map column
        /// If selected column is depth/XYZ distance unit will be visible. 
        /// </summary>
        public void SetDistanceUnitVisibility(bool isDepthColumnSelected)
        {
            ColumnViewModel depthColumn = this.ColumnsView.Where(columnValue => columnValue.SelectedWWTColumn.IsDepthColumn()).FirstOrDefault();
            ColumnViewModel xyzColumn = this.ColumnsView.Where(columnValue => columnValue.SelectedWWTColumn.IsXYZColumn()).FirstOrDefault();

            if (depthColumn != null || xyzColumn != null)
            {
                this.IsDistanceVisible = true;

                if (isDepthColumnSelected)
                {
                    if (depthColumn != null)
                    {
                        switch (depthColumn.SelectedWWTColumn.ColType)
                        {
                            case ColumnType.Alt:
                            case ColumnType.Distance:
                                this.SelectedDistanceUnit = this.DistanceUnits.Where(distanceUnit => distanceUnit.Key == AltUnit.Meters).FirstOrDefault();
                                break;

                            case ColumnType.Depth:
                                this.SelectedDistanceUnit = this.DistanceUnits.Where(distanceUnit => distanceUnit.Key == AltUnit.Kilometers).FirstOrDefault();
                                break;
                        }
                    }
                    else
                    {
                        this.SelectedDistanceUnit = this.DistanceUnits.Where(distanceUnit => distanceUnit.Key == AltUnit.Meters).FirstOrDefault();
                    }
                }
            }
            else
            {
                this.IsDistanceVisible = false;
            }
        }

        /// <summary>
        /// Sets RA unit visibility for the selected column
        /// </summary>
        public void SetRAUnitVisibility()
        {
            ColumnViewModel columnView = this.ColumnsView.Where(column => column.SelectedWWTColumn.IsRAColumn()).FirstOrDefault();
            if (columnView != null)
            {
                this.SelectedRAUnit = this.RightAscentionUnits.Where(raUnit => raUnit.Key == this.Currentlayer.LayerDetails.RAUnit).FirstOrDefault();
                this.IsRAUnitVisible = true;
            }
            else
            {
                this.SelectedRAUnit = this.RightAscentionUnits.Where(raUnit => raUnit.Key == AngleUnit.Hours).FirstOrDefault();
                this.IsRAUnitVisible = false;
            }
        }

        /// <summary>
        /// Sets marker tab visibility based on mapped columns
        /// </summary>
        public void SetMarkerTabVisibility()
        {
            this.IsMarkerTabEnabled = true;
        }

        /// <summary>
        /// Gets layer name with the proper suffix based on current layer map type
        /// 1. Local - "not in synch"
        /// 2. Local in WWT and Not in synch - "not in synch"
        /// 3. Local in WWT and in synch - "in synch"
        /// </summary>
        /// <param name="layerMap">Selected layer map</param>
        /// <param name="layerNameValue">Selected layer name</param>
        /// <returns>Layer name</returns>
        internal static string GetLayerNameOnMapType(LayerMap layerMap, string layerNameValue)
        {
            string layerName = layerNameValue;
            switch (layerMap.MapType)
            {
                case LayerMapType.Local:
                    layerName += Properties.Resources.LayerLocalText;
                    break;
                case LayerMapType.LocalInWWT:
                    if (layerMap.IsNotInSync)
                    {
                        layerName += Properties.Resources.LayerLocalText;
                    }
                    else if (!layerMap.IsNotInSync)
                    {
                        layerName += Properties.Resources.LayerLocalInWWTText;
                    }
                    break;
                default:
                    break;
            }
            return layerName;
        }

        /// <summary>
        /// Populates columns with the view model column 
        /// property and rebinds the columns
        /// </summary>
        /// <param name="selectedLayerMap">Selected layer map</param>
        /// <param name="columns">Collection of columns</param>
        internal void PopulateColumns(LayerMap selectedLayerMap, Collection<Column> columns)
        {
            this.ColumnsView = new ObservableCollection<ColumnViewModel>();
            int index = 0;

            foreach (string headerData in selectedLayerMap.HeaderRowData)
            {
                ColumnViewModel columnViewModel = new ColumnViewModel();
                columnViewModel.ExcelHeaderColumn = headerData;
                columnViewModel.WWTColumns = new ObservableCollection<Column>();
                columns.ToList().ForEach(col => columnViewModel.WWTColumns.Add(col));   
                columnViewModel.SelectedWWTColumn = columns.Where(column => column.ColType == selectedLayerMap.MappedColumnType[index]).FirstOrDefault() ?? columns.Where(column => column.ColType == ColumnType.None).FirstOrDefault();
                this.ColumnsView.Add(columnViewModel);  
                index++;
            }
            SetRAUnitVisibility();
        }

        /// <summary>
        /// Sets selected layer name value
        /// </summary>
        /// <param name="layerName">Layer name</param>
        internal void SetSelectedLayerValues(string layerName)
        {
            WorkflowController.Instance.RebuildGroupLayerDropdown();
            this.SelectedLayerText = GetLayerNameOnMapType(this.currentLayer, layerName);
        }

        /// <summary>
        /// Starts the timer for call out visibility.
        /// </summary>
        internal void StartShowHighlightAnimationTimer()
        {
            WorkflowController.Instance.BeginShowHighlightAnimation();
            DispatcherTimer myDispatcherTimer = new DispatcherTimer();
            myDispatcherTimer.Interval = TimeSpan.FromSeconds(Common.Constants.ShowHighlightCalloutTimerInterval);
            myDispatcherTimer.Tick += new EventHandler(OnShowHighlightAnimation);
            myDispatcherTimer.Start();
        }
        #endregion

        #region Events

        /// <summary>
        /// Event is fired on layer collection changed 
        /// </summary>
        /// <param name="sender">Layer collection</param>
        /// <param name="e">Routed event</param>
        protected void LayersCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e != null)
            {
                if (e.Action == NotifyCollectionChangedAction.Add)
                {
                    foreach (LayerMapDropDownViewModel layerMapModel in e.NewItems)
                    {
                        layerMapModel.LayerSelectionChangedEvent += new EventHandler(OnLayerSelectionChangedEvent);
                    }
                }
            }
        }

        /// <summary>
        /// Event is fired on Columns view collection changed
        /// </summary>
        /// <param name="sender">Map column collection</param>
        /// <param name="e">Routed event</param>
        protected void MapColumnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e != null)
            {
                if (e.Action == NotifyCollectionChangedAction.Add)
                {
                    foreach (ColumnViewModel mapColumn in e.NewItems)
                    {
                        mapColumn.MapColumnSelectionChangedEvent += new EventHandler(OnMapColumnSelectionChanged);
                    }
                }
            }
        }

        /// <summary>
        /// Event is fired when reference group collection changes
        /// </summary>
        /// <param name="sender">Reference frame</param>
        /// <param name="e">Routed event</param>
        protected void GroupCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e != null)
            {
                if (e.Action == NotifyCollectionChangedAction.Add)
                {
                    foreach (GroupViewModel groupViewModel in e.NewItems)
                    {
                        groupViewModel.GroupSelectionChangedEvent += new EventHandler(OnGroupSelectionChanged);
                    }
                }
            }
        }

        /// <summary>
        /// Event is fired on the group selection changed
        /// </summary>
        /// <param name="sender">Reference frame dropdown</param>
        /// <param name="e">Routed event</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        protected void OnGroupSelectionChanged(object sender, EventArgs e)
        {
            Group group = sender as Group;
            if (group != null)
            {
                try
                {
                    this.SetRADECAutoMap(group);

                    this.SelectedGroupText = group.Name;
                    this.SelectedGroup = group;

                    this.SetSelectedScaleType();

                    // Binding the column data to the map columns
                    Collection<Column> columns = ColumnExtensions.PopulateColumnList();

                    // Remove the columns based on the group selected.
                    this.RemoveColumns(columns);

                    this.PopulateColumns(this.currentLayer, columns);

                    // Raises event to set the object model properties and save workbook
                    this.OnCustomTaskPaneStateChanged();
                }
                catch (Exception exception)
                {
                    Logger.LogException(exception);
                }
            }
        }

        /// <summary>
        /// Event is fired on the layer selection changed
        /// </summary>
        /// <param name="sender">Layer/reference frame/layer group drop down</param>
        /// <param name="e">Routed event arguments</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        protected void OnLayerSelectionChangedEvent(object sender, EventArgs e)
        {
            LayerMap selectedLayerValue = sender as LayerMap;
            if (selectedLayerValue != null)
            {
                try
                {
                    this.SelectedLayerText = selectedLayerValue.LayerDetails.Name;
                    if (!LayerDetailsViewModel.IsPropertyChangedFromCode)
                    {
                        this.LayerSelectionChangedEvent.OnFire(selectedLayerValue, new EventArgs());
                    }
                }
                catch (Exception exception)
                {
                    Logger.LogException(exception);
                }
            }
        }

        /// <summary>
        /// Event is fired on map column selection changed
        /// from the dropdown
        /// </summary>
        /// <param name="sender">Map column</param>
        /// <param name="e">Routed event</param>
        protected void OnMapColumnSelectionChanged(object sender, EventArgs e)
        {
            if (!LayerDetailsViewModel.IsPropertyChangedFromCode)
            {
                ColumnViewModel columnViewModel = sender as ColumnViewModel;
                if (columnViewModel != null)
                {
                    var selectedColumnIndex = this.ColumnsView.IndexOf(columnViewModel);
                    if (columnViewModel.SelectedWWTColumn.ColType != ColumnType.None)
                    {
                        ColumnViewModel alreadySelectedColumn = this.ColumnsView.Where(columnValue => columnValue.SelectedWWTColumn.ColType == columnViewModel.SelectedWWTColumn.ColType
                            && selectedColumnIndex != this.ColumnsView.IndexOf(columnValue)).FirstOrDefault();

                        if (alreadySelectedColumn != null)
                        {
                            alreadySelectedColumn.SelectedWWTColumn = alreadySelectedColumn.WWTColumns[0];
                        }
                        else if (columnViewModel.SelectedWWTColumn.IsDepthColumn())
                        {
                            // If the selected column is depth column and is not currently selected
                            alreadySelectedColumn = this.ColumnsView.Where(columnValue => columnValue.SelectedWWTColumn.ColType != columnViewModel.SelectedWWTColumn.ColType
                            && columnValue.SelectedWWTColumn.IsDepthColumn()).FirstOrDefault();
                            if (alreadySelectedColumn != null)
                            {
                                alreadySelectedColumn.SelectedWWTColumn = alreadySelectedColumn.WWTColumns[0];
                            }
                        }

                        // Validates if one of X Or ReverseX/ Y or ReverseY/ Z or ReverseZ is mapped
                        ValidateXYZMapping(columnViewModel);

                        // If depth/alt is selected xyz columns will be set to "Select One" else if XYZ is selected
                        // depth/alt column is set to "Select One"
                        ValidateDepthXYZColumns(columnViewModel);
                    }
                }

                // Set Size and Name column drop downs as they very depending on the binding.
                // This needs to be done without raising an event
                LayerDetailsViewModel.IsPropertyChangedFromCode = true;

                // If magnitude column is mapped, use index of magnitude column, else use index of Depth column, else use select one (-1)
                var magColumn = this.ColumnsView.Where(item => item.SelectedWWTColumn.ColType == ColumnType.Mag).FirstOrDefault();

                if (magColumn != null && magColumn.SelectedWWTColumn.ColType != ColumnType.None)
                {
                    this.SelectedSize = this.SizeColumnList[this.ColumnsView.IndexOf(magColumn) + 1];
                }
                else
                {
                    this.SelectedSize = this.SizeColumnList[0];
                }

                // Adds or removes depth/altitude column from column mapping
                // If x or y or z is mapped and no lat/long is mapped depth/altitude is removed 
                // else if lat/long is mapped depth and altitude is added back to the list.
                AddRemoveDepthAlt();

                // Sets the scale factor to 1 if magnitude is mapped in size column else set scale factor to 8
                SetColumnMapping(this.SelectedSize);

                LayerDetailsViewModel.IsPropertyChangedFromCode = false;

                this.SetMarkerTabVisibility();

                // Sets the RA Unit visibility
                this.SetRAUnitVisibility();

                this.SetDistanceUnitVisibility(columnViewModel.SelectedWWTColumn.IsDepthColumn() || columnViewModel.SelectedWWTColumn.IsXYZColumn());

                ////Raises event to set the object model properties and save workbook
                this.OnCustomTaskPaneStateChanged();
            }
        }

        /// <summary>
        /// Event is fired on call out timer complete
        /// </summary>
        /// <param name="sender">Dispatcher timer</param>
        /// <param name="e">Routed event</param>
        protected void OnCallOutTimerComplete(object sender, EventArgs e)
        {
            DispatcherTimer myDispatcherTimer = sender as DispatcherTimer;
            this.IsCallOutVisible = false;
            myDispatcherTimer.Stop();
            myDispatcherTimer = null;
        }

        /// <summary>
        /// Event is fired on show highlight animation
        /// </summary>
        /// <param name="sender">Dispatcher timer</param>
        /// <param name="e">Routed event</param>
        protected void OnShowHighlightAnimation(object sender, EventArgs e)
        {
            DispatcherTimer myDispatcherTimer = sender as DispatcherTimer;
            myDispatcherTimer.Stop();
            myDispatcherTimer = null;
            WorkflowController.Instance.BeginHideHighlightAnimation();
        }

        #endregion

        #region Private Methods

        private static Collection<double> BuildLayerOpacity()
        {
            Collection<double> opacityList = new Collection<double>();
            for (int i = 0; i <= 100; i++)
            {
                opacityList.Add(i);
            }

            return opacityList;
        }

        private static Collection<KeyValuePair<FadeType, string>> PopulateFadeType()
        {
            Collection<KeyValuePair<FadeType, string>> fadetypeValues = new Collection<KeyValuePair<FadeType, string>>();
            fadetypeValues.Add(new KeyValuePair<FadeType, string>(FadeType.None, Resources.FadeNone));
            fadetypeValues.Add(new KeyValuePair<FadeType, string>(FadeType.In, Resources.FadeIn));
            fadetypeValues.Add(new KeyValuePair<FadeType, string>(FadeType.Out, Resources.FadeOut));
            fadetypeValues.Add(new KeyValuePair<FadeType, string>(FadeType.Both, Resources.FadeBoth));
            return fadetypeValues;
        }

        private static Collection<KeyValuePair<ScaleType, string>> PopulateScaleType()
        {
            Collection<KeyValuePair<ScaleType, string>> scaleTypeValues = new Collection<KeyValuePair<ScaleType, string>>();
            scaleTypeValues.Add(new KeyValuePair<ScaleType, string>(ScaleType.Power, Resources.ScaleTypePower));
            scaleTypeValues.Add(new KeyValuePair<ScaleType, string>(ScaleType.Constant, Resources.ScaleTypeConstant));
            scaleTypeValues.Add(new KeyValuePair<ScaleType, string>(ScaleType.Linear, Resources.ScaleTypeLinear));
            scaleTypeValues.Add(new KeyValuePair<ScaleType, string>(ScaleType.Log, Resources.ScaleTypeLog));
            scaleTypeValues.Add(new KeyValuePair<ScaleType, string>(ScaleType.StellarMagnitude, Resources.ScaleTypeStellarMagnitude));
            return scaleTypeValues;
        }

        /// <summary>
        /// Populates marker types
        /// </summary>
        /// <returns>Collections of Marker types and title</returns>
        private static Collection<KeyValuePair<MarkerType, string>> PopulateMarkerType()
        {
            Collection<KeyValuePair<MarkerType, string>> markerTypeValues = new Collection<KeyValuePair<MarkerType, string>>();
            MarkerType[] markers = (MarkerType[])Enum.GetValues(typeof(MarkerType));
            foreach (var markerValue in markers)
            {
                markerTypeValues.Add(new KeyValuePair<MarkerType, string>(markerValue, markerValue.ToString()));
            }
            return markerTypeValues;
        }

        /// <summary>
        /// Populates push pin types
        /// </summary>
        /// <returns>Collections of Pushpins</returns>
        private static Collection<KeyValuePair<int, BitmapImage>> PopulatePushpinType()
        {
            Collection<KeyValuePair<int, BitmapImage>> pushPins = new Collection<KeyValuePair<int, BitmapImage>>();

            for (int i = 0; i < PushPin.PinCount; i++)
            {
                pushPins.Add(new KeyValuePair<int, BitmapImage>(i, PushPin.GetPushPinBitmapImage(i)));
            }

            return pushPins;
        }

        private static Collection<double> BuildFactor()
        {
            Collection<double> decayList = new Collection<double>();
            foreach (int value in ComputePower(2, 12).Reverse())
            {
                decayList.Add((double)1 / (int)value);
            }
            decayList.Add(1);
            foreach (int value in ComputePower(2, 12))
            {
                decayList.Add((int)value);
            }

            return decayList;
        }

        private static IEnumerable<int> ComputePower(int number, int exponent)
        {
            int exponentNum = 0;
            int numberResult = 1;
            while (exponentNum < exponent)
            {
                numberResult *= number;
                exponentNum++;
                yield return numberResult;
            }
        }

        private static Collection<KeyValuePair<ScaleRelativeType, string>> PopulateScaleRelatives()
        {
            Collection<KeyValuePair<ScaleRelativeType, string>> scaleRelativeValues = new Collection<KeyValuePair<ScaleRelativeType, string>>();
            scaleRelativeValues.Add(new KeyValuePair<ScaleRelativeType, string>(ScaleRelativeType.World, Resources.ScaleRelativeWorld));
            scaleRelativeValues.Add(new KeyValuePair<ScaleRelativeType, string>(ScaleRelativeType.Screen, Resources.ScaleRelativeScreen));
            return scaleRelativeValues;
        }

        private static Collection<KeyValuePair<AltUnit, string>> PopulateDistanceUnits()
        {
            Collection<KeyValuePair<AltUnit, string>> distanceUnitValues = new Collection<KeyValuePair<AltUnit, string>>();
            distanceUnitValues.Add(new KeyValuePair<AltUnit, string>(AltUnit.Inches, Resources.DistanceInches));
            distanceUnitValues.Add(new KeyValuePair<AltUnit, string>(AltUnit.Feet, Resources.DistanceFeet));
            distanceUnitValues.Add(new KeyValuePair<AltUnit, string>(AltUnit.Miles, Resources.DistanceMiles));
            distanceUnitValues.Add(new KeyValuePair<AltUnit, string>(AltUnit.Meters, Resources.DistanceMeters));
            distanceUnitValues.Add(new KeyValuePair<AltUnit, string>(AltUnit.Kilometers, Resources.DistanceKiloMeters));
            distanceUnitValues.Add(new KeyValuePair<AltUnit, string>(AltUnit.AstronomicalUnits, Resources.DistanceAstronomicalUnits));
            distanceUnitValues.Add(new KeyValuePair<AltUnit, string>(AltUnit.LightYears, Resources.DistanceLightYears));
            distanceUnitValues.Add(new KeyValuePair<AltUnit, string>(AltUnit.Parsecs, Resources.DistanceParsecs));
            distanceUnitValues.Add(new KeyValuePair<AltUnit, string>(AltUnit.MegaParsecs, Resources.DistanceMegaParsecs));
            return distanceUnitValues;
        }

        /// <summary>
        /// Populates RA Units 
        /// </summary>
        /// <returns>Collection of static RA units</returns>
        private static Collection<KeyValuePair<AngleUnit, string>> PopulateRAUnits()
        {
            Collection<KeyValuePair<AngleUnit, string>> rightAscentionUnitValues = new Collection<KeyValuePair<AngleUnit, string>>();
            rightAscentionUnitValues.Add(new KeyValuePair<AngleUnit, string>(AngleUnit.Hours, Resources.RAHour));
            rightAscentionUnitValues.Add(new KeyValuePair<AngleUnit, string>(AngleUnit.Degrees, Resources.RADegree));
            return rightAscentionUnitValues;
        }

        /// <summary>
        /// Converts color to string
        /// </summary>
        /// <param name="colorBrush">Color brush</param>
        /// <returns>String value for the color brush</returns>
        private static string ConvertColorToString(Brush colorBrush)
        {
            string colorBrushValue = string.Empty;
            if (colorBrush != null)
            {
                SolidColorBrush color = (SolidColorBrush)colorBrush;
                StringBuilder colorValue = new StringBuilder();
                colorValue.Append(Common.Constants.ColorPrefix);
                colorValue.Append(color.Color.A.ToString(CultureInfo.InvariantCulture) + Common.Constants.ColorSeparator);
                colorValue.Append(color.Color.R.ToString(CultureInfo.InvariantCulture) + Common.Constants.ColorSeparator);
                colorValue.Append(color.Color.G.ToString(CultureInfo.InvariantCulture) + Common.Constants.ColorSeparator);
                colorValue.Append(color.Color.B.ToString(CultureInfo.InvariantCulture));
                colorBrushValue = colorValue.ToString();
            }

            return colorBrushValue;
        }

        /// <summary>
        /// Validates if start date is less than end date
        /// </summary>
        /// <param name="startDate">Start date time</param>
        /// <param name="endDate">End date time</param>
        /// <returns>If the dates are valid</returns>
        private static bool ValidateDate(DateTime startDate, DateTime endDate)
        {
            return (DateTime.Compare(startDate, endDate) <= 0);
        }

        /// <summary>
        /// Sets layer name in the dropdown and change the layer name 
        /// in the current workbook
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        private void SetLayerName(string layerName)
        {
            try
            {
                if (!LayerDetailsViewModel.IsPropertyChangedFromCode)
                {
                    this.SetSelectedLayerValues(layerName);

                    this.OnCustomTaskPaneStateChanged();
                }
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
            }
        }

        /// <summary>
        /// Raises an event to workflow controller to set auto map if the selection
        /// is changed from Sky to planet or vice-versa
        /// </summary>
        /// <param name="selectedGroupValue">Selected group value</param>
        private void SetRADECAutoMap(Group selectedGroupValue)
        {
            if (selectedGroupValue.IsPlanet() != this.SelectedGroup.IsPlanet())
            {
                this.ReferenceSelectionChanged.OnFire(this, new EventArgs());
            }
        }

        /// <summary>
        /// Sets the scale type to stellar magnitude for planet reference frame
        /// </summary>
        private void SetSelectedScaleType()
        {
            LayerDetailsViewModel.IsPropertyChangedFromCode = true;

            if (!this.SelectedGroup.IsPlanet())
            {
                this.SelectedScaleType = this.ScaleTypes.Where(scaleType => scaleType.Key == ScaleType.StellarMagnitude).FirstOrDefault();
            }
            else
            {
                if (this.SelectedScaleType.Key == ScaleType.StellarMagnitude)
                {
                    this.SelectedScaleType = this.ScaleTypes.Where(scaleType => scaleType.Key == ScaleType.Power).FirstOrDefault();
                }
            }

            LayerDetailsViewModel.IsPropertyChangedFromCode = false;
        }

        private void BindDatatoViewModel()
        {
            if (this.currentLayer != null)
            {
                this.IsTabVisible = true;
                this.IsHelpTextVisible = false;
                this.SetLayerMap(this.Currentlayer);
            }
            else
            {
                this.IsTabVisible = false;
                this.IsHelpTextVisible = true;
            }
        }

        /// <summary>
        /// Attaches command to command handler
        /// </summary>
        private void AttachCommands()
        {
            this.selectionCommand = new LayerSelectionHandler(this);
            this.controlCommand = new ControlStateChangeHandler(this);
            this.viewInWWTCommand = new ViewInWWTHandler(this);
            this.colorPalletCommand = new ColorPalletHandler(this);
            this.callOutCommand = new CallOutHandler(this);
            this.showRangeCommand = new ShowRangeHandler(this);
            this.deleteMappingCommand = new DeleteMappingHandler(this);
            this.getLayerDataCommand = new GetLayerDataHandler(this);
            this.updateLayerCommand = new UpdateLayerHandler(this);
            this.layerMapNameChangeCommand = new LayerMapNameChangeHandler(this);
            this.fadeTimeChangeCommand = new FadeTimeChangeHandler(this);
            this.refreshDropDownCommand = new RefreshDropDownHandler(this);
            this.refreshGroupDropDownCommand = new RefreshGroupDropDownHandler(this);
            this.sizeColumnChangeCommand = new SizeColumnChangehandler(this);
            this.downloadUpdatesCommand = new DownloadUpdatesHandler(this);
        }

        private void SetDefaultValues()
        {
            this.fadetypes = PopulateFadeType();
            this.scaleTypes = PopulateScaleType();
            this.markerTypes = PopulateMarkerType();
            this.pushpinTypes = PopulatePushpinType();
            this.timeDecay = new SliderViewModel(BuildFactor());
            this.layerOpacity = new SliderViewModel(BuildLayerOpacity());
            this.scaleFactor = new SliderViewModel(BuildFactor());
            this.scaleRelatives = PopulateScaleRelatives();
            this.distanceUnits = PopulateDistanceUnits();
            this.rightAscentionUnits = PopulateRAUnits();

            this.scaleFactor.SelectedSliderValue = GetSelectedScaleFactor(1);
            this.selectedDistanceUnit = this.distanceUnits[4];
            this.selectedFadeType = this.fadetypes[0];
            this.selectedScaleType = this.scaleTypes[0];
            this.selectedScaleRelative = this.scaleRelatives[0];
            this.selectedRAUnit = this.rightAscentionUnits[0];
            this.selectedMarkerType = this.markerTypes[0];

            SetDefaultBackground();

            this.IsViewInWWTEnabled = true;
            this.IsLayerInSyncInfoVisible = false;
            this.IsMarkerTabEnabled = true;
            this.IsDistanceVisible = false;
        }

        private void SetDefaultBackground()
        {
            SolidColorBrush color = new SolidColorBrush(System.Windows.Media.Color.FromArgb(System.Drawing.Color.Red.A, System.Drawing.Color.Red.R, System.Drawing.Color.Red.G, System.Drawing.Color.Red.B));
            this.ColorBackground = color;
        }

        /// <summary>
        /// Starts the timer for call out visibility.
        /// </summary>
        private void StartCallOutVisibilityTimer()
        {
            DispatcherTimer myDispatcherTimer = new DispatcherTimer();
            myDispatcherTimer.Interval = TimeSpan.FromSeconds(Common.Constants.CalloutTimerInterval);
            myDispatcherTimer.Tick += new EventHandler(OnCallOutTimerComplete);
            myDispatcherTimer.Start();
        }

        /// <summary>
        /// Sets the mapped column in 
        /// If size column is set to a column which doesn't have any mapping, the mapping for
        /// column is set to "Magnitude"
        /// </summary>
        /// <param name="selectedSize">Selected size value pair</param>
        private void SetColumnMapping(KeyValuePair<int, string> selectedSize)
        {
            if (selectedSize.Key != -1)
            {
                foreach (ColumnViewModel columnView in this.ColumnsView)
                {
                    if (columnView.ExcelHeaderColumn.Equals(selectedSize.Value, StringComparison.Ordinal) && columnView.SelectedWWTColumn.ColType == ColumnType.None)
                    {
                        columnView.SelectedWWTColumn = columnView.WWTColumns.Where(columnValue => columnValue.ColType == ColumnType.Mag).FirstOrDefault();
                    }
                }

                this.ScaleFactor.SelectedSliderValue = GetSelectedScaleFactor(1);
            }
            else
            {
                this.ScaleFactor.SelectedSliderValue = GetSelectedScaleFactor(Constants.DefaultScaleFactor);
            }

            if (!LayerDetailsViewModel.IsPropertyChangedFromCode)
            {
                this.OnCustomTaskPaneStateChanged();
            }
        }

        /// <summary>
        /// Gets the selected scale factor and converts it to the scale factor to be 
        /// shown in the UI.
        /// </summary>
        /// <param name="selectedScaleFactor">Selected scale factor</param>
        /// <returns>Returns selected scale factor</returns>
        private double GetSelectedScaleFactor(double selectedScaleFactor)
        {
            return scaleFactor.SliderTicks.IndexOf(selectedScaleFactor) + 1;
        }

        /// <summary>
        /// Validates if the collection of columns contains depth/altitude columns
        /// </summary>
        /// <returns>True if the depth and altitude column is present</returns>
        private bool ValidateDepthAltColumn()
        {
            bool isDepthCol = true;
            foreach (ColumnViewModel colVal in this.ColumnsView)
            {
                if (!colVal.WWTColumns.Contains(colVal.WWTColumns.Where(wwtCol => wwtCol.ColType == ColumnType.Depth || wwtCol.ColType == ColumnType.Alt || wwtCol.ColType == ColumnType.Distance).FirstOrDefault()))
                {
                    isDepthCol = false;
                    break;
                }
            }
            return isDepthCol;
        }

        /// <summary>
        /// Adds or removes depth/altitude column from column mapping
        /// If x or y or z is mapped and no lat/long is mapped depth/altitude is removed 
        /// else if lat/long is mapped depth and altitude is added back to the list.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity", Justification = "Adding/removing Alt columns should not separated out.")]
        private void AddRemoveDepthAlt()
        {
            var xyzCol = this.ColumnsView.Where(item => item.SelectedWWTColumn.IsXYZColumn()).FirstOrDefault();
            var latLongRADec = this.ColumnsView.Where(item => item.SelectedWWTColumn.ColType == ColumnType.Lat || item.SelectedWWTColumn.ColType == ColumnType.Long ||
                item.SelectedWWTColumn.ColType == ColumnType.RA || item.SelectedWWTColumn.ColType == ColumnType.Dec).FirstOrDefault();

            if (xyzCol != null && latLongRADec == null)
            {
                this.ColumnsView.ToList().ForEach(colVal =>
                {
                    if (colVal.SelectedWWTColumn.ColType == ColumnType.Depth || colVal.SelectedWWTColumn.ColType == ColumnType.Alt || colVal.SelectedWWTColumn.ColType == ColumnType.Distance)
                    {
                        colVal.SelectedWWTColumn = colVal.WWTColumns.Where(col => col.ColType == ColumnType.None).FirstOrDefault();
                    }
                    colVal.WWTColumns.Remove(colVal.WWTColumns.Where(wwtCol => wwtCol.ColType == ColumnType.Depth).FirstOrDefault());
                    colVal.WWTColumns.Remove(colVal.WWTColumns.Where(wwtCol => wwtCol.ColType == ColumnType.Alt).FirstOrDefault());
                    colVal.WWTColumns.Remove(colVal.WWTColumns.Where(wwtCol => wwtCol.ColType == ColumnType.Distance).FirstOrDefault());
                });
            }
            else if (!ValidateDepthAltColumn())
            {
                this.ColumnsView.ToList().ForEach(colVal =>
                {
                    colVal.WWTColumns.Add(ColumnExtensions.GetDepthColumn());
                    colVal.WWTColumns.Add(ColumnExtensions.GetAltColumn());
                    colVal.WWTColumns.Add(ColumnExtensions.GetDistanceColumn());
                });
            }
        }

        /// <summary>
        /// Validates the mappings for XYZ column.
        /// Only one of X or ReverseX / Y or ReverseY/ Z or ReverseZ can be mapped at a time.
        /// </summary>
        /// <param name="columnViewModel">Selected column</param>
        private void ValidateXYZMapping(ColumnViewModel columnViewModel)
        {
            // If the selected column in the mapped columns is X/Y/Z 
            if (columnViewModel.SelectedWWTColumn.IsXYZColumn())
            {
                ColumnViewModel xyzSelectedColumn = null;
                switch (columnViewModel.SelectedWWTColumn.ColType)
                {
                    case ColumnType.X:
                        xyzSelectedColumn = this.ColumnsView.Where(columnValue => columnValue.SelectedWWTColumn.ColType == ColumnType.ReverseX).FirstOrDefault();
                        break;
                    case ColumnType.Y:
                        xyzSelectedColumn = this.ColumnsView.Where(columnValue => columnValue.SelectedWWTColumn.ColType == ColumnType.ReverseY).FirstOrDefault();
                        break;
                    case ColumnType.Z:
                        xyzSelectedColumn = this.ColumnsView.Where(columnValue => columnValue.SelectedWWTColumn.ColType == ColumnType.ReverseZ).FirstOrDefault();
                        break;
                    case ColumnType.ReverseX:
                        xyzSelectedColumn = this.ColumnsView.Where(columnValue => columnValue.SelectedWWTColumn.ColType == ColumnType.X).FirstOrDefault();
                        break;
                    case ColumnType.ReverseY:
                        xyzSelectedColumn = this.ColumnsView.Where(columnValue => columnValue.SelectedWWTColumn.ColType == ColumnType.Y).FirstOrDefault();
                        break;
                    case ColumnType.ReverseZ:
                        xyzSelectedColumn = this.ColumnsView.Where(columnValue => columnValue.SelectedWWTColumn.ColType == ColumnType.Z).FirstOrDefault();
                        break;
                    default:
                        break;
                }
                if (xyzSelectedColumn != null)
                {
                    xyzSelectedColumn.SelectedWWTColumn = xyzSelectedColumn.WWTColumns[0];
                }
            }
        }

        /// <summary>
        /// Validates if depth column is mapped and xyz columns are also mapped,
        /// If depth/alt is selected xyz columns will be set to "Select One" else if XYZ is selected
        /// depth/alt column is set to "Select One"
        /// </summary>
        /// <param name="columnViewModel">Selected column in mapping</param>
        private void ValidateDepthXYZColumns(ColumnViewModel columnViewModel)
        {
            // If the selected column is depth column and there is already xyz column mapped then 
            // all the xyz selection would be changed to "Select One"
            if (columnViewModel.SelectedWWTColumn.IsDepthColumn() && !this.currentLayer.IsXYZLayer())
            {
                ColumnViewModel xyzColumn = null;
                xyzColumn = this.ColumnsView.Where(columnValue => columnValue.SelectedWWTColumn.ColType == ColumnType.ReverseX || columnValue.SelectedWWTColumn.ColType == ColumnType.X).FirstOrDefault();
                if (xyzColumn != null)
                {
                    xyzColumn.SelectedWWTColumn = xyzColumn.WWTColumns[0];
                }
                xyzColumn = this.ColumnsView.Where(columnValue => columnValue.SelectedWWTColumn.ColType == ColumnType.ReverseY || columnValue.SelectedWWTColumn.ColType == ColumnType.Y).FirstOrDefault();
                if (xyzColumn != null)
                {
                    xyzColumn.SelectedWWTColumn = xyzColumn.WWTColumns[0];
                }
                xyzColumn = this.ColumnsView.Where(columnValue => columnValue.SelectedWWTColumn.ColType == ColumnType.ReverseZ || columnValue.SelectedWWTColumn.ColType == ColumnType.Z).FirstOrDefault();
                if (xyzColumn != null)
                {
                    xyzColumn.SelectedWWTColumn = xyzColumn.WWTColumns[0];
                }
            }
            else if (columnViewModel.SelectedWWTColumn.IsXYZColumn() && this.ColumnsView.Where(colVal => colVal.SelectedWWTColumn.IsDepthColumn()).Any())
            {
                // If selected column is xyz column and depth column was already mapped , the depth column 
                // would be set to "Select One"
                ColumnViewModel depthCol = this.ColumnsView.Where(colVal => colVal.SelectedWWTColumn.IsDepthColumn()).FirstOrDefault();
                if (depthCol != null)
                {
                    depthCol.SelectedWWTColumn = depthCol.WWTColumns[0];
                }
            }
        }

        #endregion

        #region Event Handler

        private class LayerSelectionHandler : RelayCommand
        {
            private LayerDetailsViewModel parent;

            public LayerSelectionHandler(LayerDetailsViewModel layerDetailsViewModel)
            {
                this.parent = layerDetailsViewModel;
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
            public override void Execute(object parameter)
            {
                LayerMapDropDownViewModel layerMapModel = parameter as LayerMapDropDownViewModel;
                if (layerMapModel != null)
                {
                    try
                    {
                        if (layerMapModel.Name.Equals(Resources.DefaultSelectedLayerName, StringComparison.OrdinalIgnoreCase))
                        {
                            this.parent.SelectedLayerText = layerMapModel.Name;
                            this.parent.Currentlayer = null;
                            if (!LayerDetailsViewModel.IsPropertyChangedFromCode)
                            {
                                this.parent.LayerSelectionChangedEvent.OnFire(this.parent.currentLayer, new EventArgs());
                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        Logger.LogException(exception);
                    }
                }
            }
        }

        private class ControlStateChangeHandler : RelayCommand
        {
            private LayerDetailsViewModel parent;

            public ControlStateChangeHandler(LayerDetailsViewModel layerDetailsViewModel)
            {
                this.parent = layerDetailsViewModel;
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
            public override void Execute(object parameter)
            {
                try
                {
                    if (!LayerDetailsViewModel.IsPropertyChangedFromCode && this.parent != null)
                    {
                        this.parent.OnCustomTaskPaneStateChanged();
                    }
                }
                catch (Exception exception)
                {
                    Logger.LogException(exception);
                }
            }
        }

        private class ViewInWWTHandler : RelayCommand
        {
            private LayerDetailsViewModel parent;
            public ViewInWWTHandler(LayerDetailsViewModel layerDetailsViewModel)
            {
                this.parent = layerDetailsViewModel;
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
            public override void Execute(object parameter)
            {
                if (this.parent != null)
                {
                    try
                    {
                        this.parent.ViewnInWWTClickedEvent.OnFire(this.parent, new EventArgs());
                    }
                    catch (Exception exception)
                    {
                        Logger.LogException(exception);
                    }
                }
            }
        }

        private class ShowRangeHandler : RelayCommand
        {
            private LayerDetailsViewModel parent;
            public ShowRangeHandler(LayerDetailsViewModel layerDetailsViewModel)
            {
                this.parent = layerDetailsViewModel;
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
            public override void Execute(object parameter)
            {
                if (this.parent != null)
                {
                    try
                    {
                        this.parent.ShowRangeClickedEvent.OnFire(this.parent, new EventArgs());
                    }
                    catch (Exception exception)
                    {
                        Logger.LogException(exception);
                    }
                }
            }
        }

        private class DeleteMappingHandler : RelayCommand
        {
            private LayerDetailsViewModel parent;
            public DeleteMappingHandler(LayerDetailsViewModel layerDetailsViewModel)
            {
                this.parent = layerDetailsViewModel;
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
            public override void Execute(object parameter)
            {
                if (this.parent != null)
                {
                    try
                    {
                        this.parent.DeleteMappingClickedEvent.OnFire(this.parent, new EventArgs());
                    }
                    catch (Exception exception)
                    {
                        Logger.LogException(exception);
                    }
                }
            }
        }

        private class GetLayerDataHandler : RelayCommand
        {
            private LayerDetailsViewModel parent;
            public GetLayerDataHandler(LayerDetailsViewModel layerDetailsViewModel)
            {
                this.parent = layerDetailsViewModel;
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
            public override void Execute(object parameter)
            {
                if (this.parent != null)
                {
                    try
                    {
                        this.parent.GetLayerDataClickedEvent.OnFire(this.parent, new EventArgs());
                    }
                    catch (Exception exception)
                    {
                        Logger.LogException(exception);
                    }
                }
            }
        }

        private class ColorPalletHandler : RelayCommand
        {
            private LayerDetailsViewModel parent;
            public ColorPalletHandler(LayerDetailsViewModel layerDetailsViewModel)
            {
                this.parent = layerDetailsViewModel;
            }

            public override void Execute(object parameter)
            {
                using (ColorDialog dialog = new ColorDialog())
                {
                    dialog.Color = System.Drawing.Color.White;
                    dialog.SolidColorOnly = true;
                    dialog.AllowFullOpen = false;
                    dialog.FullOpen = false;
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        SolidColorBrush color = new SolidColorBrush(System.Windows.Media.Color.FromArgb(dialog.Color.A, dialog.Color.R, dialog.Color.G, dialog.Color.B));
                        if (this.parent != null)
                        {
                            this.parent.ColorBackground = color;
                        }
                    }
                }
            }
        }

        private class CallOutHandler : RelayCommand
        {
            private LayerDetailsViewModel parent;
            public CallOutHandler(LayerDetailsViewModel layerDetailsViewModel)
            {
                this.parent = layerDetailsViewModel;
            }

            public override void Execute(object parameter)
            {
                if (this.parent != null)
                {
                    this.parent.IsCallOutVisible = false;
                }
            }
        }

        private class UpdateLayerHandler : RelayCommand
        {
            private LayerDetailsViewModel parent;
            public UpdateLayerHandler(LayerDetailsViewModel layerDetailsViewModel)
            {
                this.parent = layerDetailsViewModel;
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
            public override void Execute(object parameter)
            {
                if (this.parent != null)
                {
                    try
                    {
                        this.parent.UpdateLayerClickedEvent.OnFire(this.parent, new EventArgs());
                    }
                    catch (Exception exception)
                    {
                        Logger.LogException(exception);
                    }
                }
            }
        }

        private class LayerMapNameChangeHandler : RelayCommand
        {
            private LayerDetailsViewModel parent;
            public LayerMapNameChangeHandler(LayerDetailsViewModel layerDetailsViewModel)
            {
                this.parent = layerDetailsViewModel;
            }

            public override void Execute(object parameter)
            {
                if (this.parent != null && parameter != null)
                {
                    string layerName = parameter as string;
                    if (layerName != null)
                    {
                        this.parent.SelectedLayerName = layerName;
                    }
                }
            }
        }

        private class FadeTimeChangeHandler : RelayCommand
        {
            private LayerDetailsViewModel parent;
            public FadeTimeChangeHandler(LayerDetailsViewModel layerDetailsViewModel)
            {
                this.parent = layerDetailsViewModel;
            }

            public override void Execute(object parameter)
            {
                if (this.parent != null && parameter != null)
                {
                    string fadeTime = parameter as string;
                    if (fadeTime != null)
                    {
                        this.parent.FadeTime = fadeTime;
                    }
                }
            }
        }

        private class RefreshDropDownHandler : RelayCommand
        {
            private LayerDetailsViewModel parent;
            public RefreshDropDownHandler(LayerDetailsViewModel layerDetailsViewModel)
            {
                this.parent = layerDetailsViewModel;
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
            public override void Execute(object parameter)
            {
                if (this.parent != null)
                {
                    try
                    {
                        this.parent.RefreshDropDownClickedEvent.OnFire(this.parent, new EventArgs());
                    }
                    catch (Exception exception)
                    {
                        Logger.LogException(exception);
                    }
                }
            }
        }

        private class RefreshGroupDropDownHandler : RelayCommand
        {
            private LayerDetailsViewModel parent;
            public RefreshGroupDropDownHandler(LayerDetailsViewModel layerDetailsViewModel)
            {
                this.parent = layerDetailsViewModel;
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
            public override void Execute(object parameter)
            {
                if (this.parent != null)
                {
                    try
                    {
                        this.parent.RefreshGroupDropDownClickedEvent.OnFire(this.parent, new EventArgs());
                    }
                    catch (Exception exception)
                    {
                        Logger.LogException(exception);
                    }
                }
            }
        }

        private class SizeColumnChangehandler : RelayCommand
        {
            private LayerDetailsViewModel parent;
            public SizeColumnChangehandler(LayerDetailsViewModel layerDetailsViewModel)
            {
                this.parent = layerDetailsViewModel;
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
            public override void Execute(object parameter)
            {
                try
                {
                    if (parameter != null && !LayerDetailsViewModel.IsPropertyChangedFromCode)
                    {
                        KeyValuePair<int, string> selectedSize = (KeyValuePair<int, string>)parameter;
                        this.parent.SetColumnMapping(selectedSize);
                    }
                }
                catch (Exception exception)
                {
                    Logger.LogException(exception);
                }
            }
        }

        private class DownloadUpdatesHandler : RelayCommand
        {
            private LayerDetailsViewModel parent;
            public DownloadUpdatesHandler(LayerDetailsViewModel layerDetailsViewModel)
            {
                this.parent = layerDetailsViewModel;
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
            public override void Execute(object parameter)
            {
                if (this.parent != null)
                {
                    try
                    {
                        this.parent.DownloadUpdatesClickedEvent.OnFire(this.parent, new EventArgs());
                    }
                    catch (Exception exception)
                    {
                        Logger.LogException(exception);
                    }
                }
            }
        }

        #endregion
    }
}
