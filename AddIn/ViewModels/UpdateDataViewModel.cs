//-----------------------------------------------------------------------
// <copyright file="UpdateDataViewModel.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.ObjectModel;
using System.Windows.Input;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// View model for update data view.
    /// </summary>
    public class UpdateDataViewModel : PropertyChangeBase
    {
        private AltitudeStyle altitudeSytle;
        private bool showAltitudeConstant;
        private bool showAltitudeLinear;

        private ColorScheme colorScheme;
        private bool showRegularScheme;
        private bool showDataScheme;
        private bool filterBetweenBoundaries;

        public double DeltaLatitude { get; set; }

        public double DeltaLongitude { get; set; }

        public ColorPallete ColorPalette { get; set; }

        public ColorScheme ColorScheme
        {
            get
            {
                return this.colorScheme;
            }
            set
            {
                this.colorScheme = value;
                switch (this.colorScheme)
                {
                    case ColorScheme.FromData:
                        this.ShowDataScheme = true;
                        this.ShowRegularScheme = false;
                        break;
                    default:
                        this.ShowDataScheme = false;
                        this.ShowRegularScheme = true;
                        break;
                }

                OnPropertyChanged("ColorScheme");
            }
        }

        public string ColorColumn { get; set; }

        public string AltitudeColumn { get; set; }

        public string RColumn { get; set; }

        public string GColumn { get; set; }

        public string BColumn { get; set; }

        public AltitudeStyle AltitudeSytle
        {
            get
            {
                return this.altitudeSytle;
            }
            set
            {
                this.altitudeSytle = value;
                switch (this.altitudeSytle)
                {
                    case AltitudeStyle.Linear:
                        this.ShowAltitudeLinear = true;
                        this.ShowAltitudeConstant = false;
                        break;
                    case AltitudeStyle.Constant:
                    default:
                        this.ShowAltitudeLinear = false;
                        this.ShowAltitudeConstant = true;
                        break;
                }

                OnPropertyChanged("AltitudeSytle");
            }
        }

        public double AltitudeConstant { get; set; }

        public double AlphaConstant { get; set; }

        public double BetaConstant { get; set; }

        public double ColorMin { get; set; }

        public double ColorMax { get; set; }
        
        public double MinLatitude { get; set; }

        public double MaxLatitude { get; set; }

        public double MinLongitude { get; set; }

        public double MaxLongitude { get; set; }

        public bool FilterBetweenBoundaries
        {
            get
            {
                return this.filterBetweenBoundaries;
            }
            set
            {
                this.filterBetweenBoundaries = value;
                OnPropertyChanged("FilterBetweenBoundaries");
            }
        }

        public bool ShowAltitudeConstant
        {
            get
            {
                return this.showAltitudeConstant;
            }
            set
            {
                this.showAltitudeConstant = value;
                OnPropertyChanged("ShowAltitudeConstant");
            }
        }

        public bool ShowAltitudeLinear
        {
            get
            {
                return this.showAltitudeLinear;
            }
            set
            {
                this.showAltitudeLinear = value;
                OnPropertyChanged("ShowAltitudeLinear");
            }
        }

        public bool ShowRegularScheme
        {
            get
            {
                return this.showRegularScheme;
            }
            set
            {
                this.showRegularScheme = value;
                OnPropertyChanged("ShowRegularScheme");
            }
        }

        public bool ShowDataScheme
        {
            get
            {
                return this.showDataScheme;
            }
            set
            {
                this.showDataScheme = value;
                OnPropertyChanged("ShowDataScheme");
            }
        }

        public static string GetGeometry(double lat, double lon, UpdateDataModel input)
        {
            //double ln1 = lon - deltaLong;
            //double lt1 = lat - deltaLat;
            //double ln2 = lon + deltaLong;
            //double lt2 = lat + deltaLat;

            // return string.Format("Polygon(({0} {1},{0} {3}, {2} {3},{2} {1},{0} {1}))", ln1, lt1, ln2, lt2);

            return string.Format(
                "Polygon(({0} {1},{0} {3}, {2} {3},{2} {1},{0} {1}))",
                lon - input.DeltaLongitude,
                lat - input.DeltaLatitude,
                lon + input.DeltaLongitude,
                lat + input.DeltaLatitude);
        }

        public static double GetAltitudeValue(double altValue, UpdateDataModel input)
        {
            double updatedValue = altValue;
            switch (input.AltitudeSytle)
            {
                case AltitudeStyle.Constant:
                    updatedValue = altValue + input.AltitudeConstant;
                    break;
                case AltitudeStyle.Linear:
                    updatedValue = input.AlphaConstant + (altValue * input.BetaConstant);
                    break;
            }

            return updatedValue;
        }

        public static string GetColorValue(int rValue, int gValue, int bValue)
        {
            return string.Format("FF{0}{1}{2}", rValue.ToString("X2"), gValue.ToString("X2"), bValue.ToString("X2"));
            // return (System.Drawing.Color.FromArgb(rValue, gValue, bValue).ToArgb() & 0x00FFFFFF).ToString("X6");
        }
    }
}
