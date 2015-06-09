//-----------------------------------------------------------------------
// <copyright file="Perspective.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Runtime.Serialization;
namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Class represents the perspective feature of WWT.
    /// </summary>
    [DataContract]
    public class Perspective
    {
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the Perspective class.
        /// </summary>
        /// <param name="lookAt">
        /// lookAt value 
        /// </param>
        /// <param name="referenceFrame">
        /// referenceFrame value 
        /// </param>
        /// <param name="hasRADec">
        /// has RA Dec values
        /// </param>
        /// <param name="latitude">
        /// Latitude/RA value 
        /// </param>
        /// <param name="longitude">
        /// Longitude/Dec value
        /// </param>
        /// <param name="zoom">
        /// Zoom level on a scale of 0 to 360
        /// with 360 representing 59200 Kilo Meters
        /// </param>
        /// <param name="rotation">
        /// Rotation in degrees.
        /// </param>
        /// <param name="lookAngle">
        /// Look angle in degrees.
        /// </param>
        /// <param name="observingTime">
        /// observingTime value
        /// </param>
        /// <param name="timeRate">
        /// timeRate value
        /// </param>
        /// <param name="zoomText">
        /// zoomText value
        /// </param>
        /// <param name="viewToken">
        /// viewToken value
        /// </param>
        public Perspective(string lookAt, string referenceFrame, bool hasRADec, string latitude, string longitude, string zoom, string rotation, string lookAngle, string observingTime, string timeRate, string zoomText, string viewToken)
        {
            this.LookAt = lookAt;
            this.ReferenceFrame = referenceFrame;
            if (hasRADec)
            {
                this.RightAscention = latitude;
                this.Declination = longitude;
            }
            else
            {
                this.Latitude = latitude;
                this.Longitude = longitude;
            }
            this.Zoom = zoom;
            this.Rotation = rotation;
            this.LookAngle = lookAngle;
            this.ObservingTime = observingTime;
            this.TimeRate = timeRate;
            this.ZoomText = zoomText;
            this.ViewToken = viewToken;
            this.HasRADec = hasRADec; 
        }

        #endregion Constructor

        #region Properties

        /// <summary>
        /// Gets or sets the Look at value 
        /// </summary>
        [DataMember]
        public string LookAt { get; set; }

        /// <summary>
        /// Gets or sets the Reference Frame value 
        /// </summary>
        [DataMember]
        public string ReferenceFrame { get; set; }

        /// <summary>
        /// Gets or sets the value of observing time
        /// </summary>
        [DataMember]
        public string ObservingTime { get; set; }

        /// <summary>
        /// Gets or sets the value of time rate
        /// </summary>
        [DataMember]
        public string TimeRate { get; set; }

        /// <summary>
        /// Gets or sets the value of latitude in degrees
        /// </summary>
        [DataMember]
        public string Latitude { get; set; }

        /// <summary>
        /// Gets or sets the value of Longitude in degrees
        /// </summary>
        [DataMember]
        public string Longitude { get; set; }

        /// <summary>
        /// Gets or sets the value of zoom level on a scale of 0 to 360
        /// with 360 representing 59200 Kilo Meters
        /// </summary>
        [DataMember]
        public string Zoom { get; set; }

        /// <summary>
        /// Gets or sets the Zoom text value 
        /// </summary>
        [DataMember]
        public string ZoomText { get; set; }

        /// <summary>
        /// Gets or sets the ViewToken value 
        /// </summary>
        [DataMember]
        public string ViewToken { get; set; }

        /// <summary>
        /// Gets or sets the value of rotation in degrees
        /// </summary>
        [DataMember]
        public string Rotation { get; set; }

        /// <summary>
        /// Gets or sets the value of look angle in degrees
        /// </summary>
        [DataMember]
        public string LookAngle { get; set; }

        /// <summary>
        /// Gets or sets the name of the perspective details
        /// </summary>
        [DataMember]
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the value of RA 
        /// </summary>
        [DataMember]
        public string RightAscention { get; set; }

        /// <summary>
        /// Gets or sets the value of Dec
        /// </summary>
        [DataMember]
        public string Declination { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether values are RA/Dec 
        /// </summary>
        [DataMember]
        public bool HasRADec { get; set; }

        #endregion Properties
    }
}
