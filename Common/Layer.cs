//-----------------------------------------------------------------------
// <copyright file="Layer.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Runtime.Serialization;

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Details of WWT layer model.
    /// </summary>
    [DataContract]
    public class Layer
    {
        /// <summary>
        /// Initializes a new instance of the Layer class.
        /// </summary>
        public Layer()
        {
            this.InitializeDefaults();
        }

        /// <summary>
        /// Gets or sets the Coordinate type of the layer.
        /// </summary>
        [DataMember]
        public CoordinatesType CoordinatesType { get; set; }

        /// <summary>
        /// Gets or sets x axis for rectangular co-ordinate type
        /// </summary>
        [DataMember]
        public int XAxis { get; set; }

        /// <summary>
        /// Gets or sets y axis for rectangular co-ordinate type
        /// </summary>
        [DataMember]
        public int YAxis { get; set; }

        /// <summary>
        /// Gets or sets z axis for rectangular co-ordinate type
        /// </summary>
        [DataMember]
        public int ZAxis { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether reverse x axis is should be checked rectangular co-ordinate type
        /// </summary>
        [DataMember]
        public bool ReverseXAxis { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether reverse y axis is should be checked rectangular co-ordinate type
        /// </summary>
        [DataMember]
        public bool ReverseYAxis { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether reverse z axis is should be checked rectangular co-ordinate type
        /// </summary>
        [DataMember]
        public bool ReverseZAxis { get; set; }

        /// <summary>
        /// Gets or sets the group of the layer.
        /// </summary>
        [DataMember]
        public Group Group { get; set; }

        /// <summary>
        /// Gets or sets the value of ID
        /// </summary>
        [DataMember]
        public string ID { get; set; }

        /// <summary>
        /// Gets or sets the value of Name
        /// </summary>
        [DataMember]
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the value of layer version
        /// </summary>
        [DataMember]
        public int Version { get; set; }

        /// <summary>
        /// Gets or sets the value of TimeDecay
        /// </summary>
        [DataMember]
        public double TimeDecay { get; set; }

        /// <summary>
        /// Gets or sets the value of ScaleFactor
        /// </summary>
        [DataMember]
        public double ScaleFactor { get; set; }

        /// <summary>
        /// Gets or sets the value of Opacity
        /// </summary>
        [DataMember]
        public double Opacity { get; set; }

        /// <summary>
        /// Gets or sets the value of StartTime
        /// </summary>
        [DataMember]
        public DateTime StartTime { get; set; }

        /// <summary>
        /// Gets or sets the value of EndTime
        /// </summary>
        [DataMember]
        public DateTime EndTime { get; set; }

        /// <summary>
        /// Gets or sets the value of FadeSpan
        /// </summary>
        [DataMember]
        public TimeSpan FadeSpan { get; set; }

        /// <summary>
        /// Gets or sets the value of Color
        /// </summary>
        [DataMember]
        public string Color { get; set; }

        /// <summary>
        /// Gets or sets the value of LatColumn
        /// </summary>
        [DataMember]
        public int LatColumn { get; set; }

        /// <summary>
        /// Gets or sets the value of LngColumn
        /// </summary>
        [DataMember]
        public int LngColumn { get; set; }

        /// <summary>
        /// Gets or sets the value of RAColumn
        /// </summary>
        [DataMember]
        public int RAColumn { get; set; }

        /// <summary>
        /// Gets or sets the value of DecColumn
        /// </summary>
        [DataMember]
        public int DecColumn { get; set; }

        /// <summary>
        /// Gets or sets the value of GeometryColumn
        /// </summary>
        [DataMember]
        public int GeometryColumn { get; set; }

        /// <summary>
        /// Gets or sets the value of ColorMapColumn
        /// </summary>
        [DataMember]
        public int ColorMapColumn { get; set; }

        /// <summary>
        /// Gets or sets the value of MarkerIndex
        /// </summary>
        [DataMember]
        public int MarkerIndex { get; set; }

        /// <summary>
        /// Gets or sets the value of plot type
        /// </summary>
        [DataMember]
        public MarkerType PlotType { get; set; }

        /// <summary>
        /// Gets or sets the value of AltColumn
        /// </summary>
        [DataMember]
        public int AltColumn { get; set; }

        /// <summary>
        /// Gets or sets the value of StartDateColumn
        /// </summary>
        [DataMember]
        public int StartDateColumn { get; set; }

        /// <summary>
        /// Gets or sets the value of EndDateColumn
        /// </summary>
        [DataMember]
        public int EndDateColumn { get; set; }

        /// <summary>
        /// Gets or sets the value of SizeColumn
        /// </summary>
        [DataMember]
        public int SizeColumn { get; set; }

        /// <summary>
        /// Gets or sets the value of NameColumn
        /// </summary>
        [DataMember]
        public int NameColumn { get; set; }

        /// <summary>
        /// Gets or sets the value of AltType
        /// </summary>
        [DataMember]
        public AltType AltType { get; set; }

        /// <summary>
        /// Gets or sets the value of MarkerScale
        /// </summary>
        [DataMember]
        public ScaleRelativeType MarkerScale { get; set; }

        /// <summary>
        /// Gets or sets the value of AltUnit
        /// </summary>
        [DataMember]
        public AltUnit AltUnit { get; set; }

        /// <summary>
        /// Gets or sets the value of RA unit
        /// </summary>
        [DataMember]
        public AngleUnit RAUnit { get; set; }

        /// <summary>
        /// Gets or sets the value of PointScaleType
        /// </summary>
        [DataMember]
        public ScaleType PointScaleType { get; set; }

        /// <summary>
        /// Gets or sets the value of FadeType
        /// </summary>
        [DataMember]
        public FadeType FadeType { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether far side is shown or not
        /// </summary>
        [DataMember]
        public bool ShowFarSide { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether time series is shown or not
        /// </summary>
        [DataMember]
        public bool HasTimeSeries { get; set; }

        /// <summary>
        /// Get default values for Layer
        /// </summary>
        public void InitializeDefaults()
        {
            this.CoordinatesType = CoordinatesType.Spherical;
            this.TimeDecay = Constants.DefaultTimeDecay;
            this.ScaleFactor = Constants.DefaultScaleFactor;
            this.Opacity = Constants.DefaultOpacity;
            this.StartTime = Constants.DefaultStartTime;
            this.EndTime = Constants.DefaultEndTime;
            this.FadeSpan = Constants.DefaultFadeSpan;
            this.Color = Constants.DefaultColor;
            this.LatColumn = Constants.DefaultColumnIndex;
            this.LngColumn = Constants.DefaultColumnIndex;
            this.GeometryColumn = Constants.DefaultColumnIndex;
            this.ColorMapColumn = Constants.DefaultColumnIndex;
            this.AltColumn = Constants.DefaultColumnIndex;
            this.StartDateColumn = Constants.DefaultColumnIndex;
            this.EndDateColumn = Constants.DefaultColumnIndex;
            this.SizeColumn = Constants.DefaultColumnIndex;
            this.NameColumn = Constants.DefaultColumnIndex;
            this.AltType = AltType.Depth;
            this.MarkerScale = ScaleRelativeType.World;
            this.AltUnit = AltUnit.Meters;
            this.RAUnit = AngleUnit.Hours;
            this.PointScaleType = ScaleType.Power;
            this.FadeType = FadeType.None;
            this.RAColumn = Constants.DefaultColumnIndex;
            this.DecColumn = Constants.DefaultColumnIndex;
            this.PlotType = MarkerType.Gaussian;
            this.MarkerIndex = Constants.DefaultMarkerIndex;
            this.XAxis = Constants.DefaultColumnIndex;
            this.YAxis = Constants.DefaultColumnIndex;
            this.ZAxis = Constants.DefaultColumnIndex;
            this.ReverseXAxis = false;
            this.ReverseYAxis = false;
            this.ReverseZAxis = false;
            this.ShowFarSide = true;
            this.HasTimeSeries = false;
            this.Version = 0;
        }

        /// <summary>
        /// Initialize default values for columns
        /// </summary>
        public void InitializeColumnDefaults()
        {
            this.LatColumn = Constants.DefaultColumnIndex;
            this.LngColumn = Constants.DefaultColumnIndex;
            this.GeometryColumn = Constants.DefaultColumnIndex;
            this.ColorMapColumn = Constants.DefaultColumnIndex;
            this.AltColumn = Constants.DefaultColumnIndex;
            this.StartDateColumn = Constants.DefaultColumnIndex;
            this.EndDateColumn = Constants.DefaultColumnIndex;
            this.SizeColumn = Constants.DefaultColumnIndex;
            this.NameColumn = Constants.DefaultColumnIndex;
            this.RAColumn = Constants.DefaultColumnIndex;
            this.DecColumn = Constants.DefaultColumnIndex;
            this.XAxis = Constants.DefaultColumnIndex;
            this.YAxis = Constants.DefaultColumnIndex;
            this.ZAxis = Constants.DefaultColumnIndex;
            this.ReverseXAxis = false;
            this.ReverseYAxis = false;
            this.ReverseZAxis = false;
        }
    }
}
