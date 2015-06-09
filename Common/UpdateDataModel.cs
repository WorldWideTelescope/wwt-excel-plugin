// -----------------------------------------------------------------------
// <copyright file="UpdateDataModel.cs">
//  View model for fetch climate view.
// </copyright>
// -----------------------------------------------------------------------

namespace Microsoft.Research.Wwt.Excel.Common
{
    using System;
    using System.Collections.Generic;
    using System.Windows.Input;
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Research.Wwt.Excel.Common;

    public enum ColorPallete
    {
        FetchClimate = 0,
        JeffDozier = 1,
        MBARI = 2
    }

    public enum ColorScheme
    {
        FromData = 1,
    }

    public enum AltitudeStyle
    {
        Constant = 0,
        Linear = 1
    }

    /// <summary>
    /// View model for update data view.
    /// </summary>
    public class UpdateDataModel
    {
        public double DeltaLatitude { get; set; }

        public double DeltaLongitude { get; set; }

        public ColorPallete ColorPalette { get; set; }

        public ColorScheme ColorScheme { get; set; }

        public string ColorColumn { get; set; }

        public string AltitudeColumn { get; set; }

        public string RColumn { get; set; }

        public string GColumn { get; set; }

        public string BColumn { get; set; }

        public AltitudeStyle AltitudeSytle { get; set; }

        public double AltitudeConstant { get; set; }

        public double AlphaConstant { get; set; }

        public double BetaConstant { get; set; }

        public double ColorMin { get; set; }

        public double ColorMax { get; set; }

        public double MinLatitude { get; set; }

        public double MaxLatitude { get; set; }

        public double MinLongitude { get; set; }

        public double MaxLongitude { get; set; }

        public bool FilterBetweenBoundaries { get; set; }

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

        public static bool CheckWithinBoundary(double lat, double lon, UpdateDataModel input)
        {
            return input.MinLongitude <= lon && lon <= input.MaxLongitude &&  input.MinLatitude <= lat && lat <= input.MaxLatitude;
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

        public static string GetCircle(double centerLat, double centerLon, double radius)
        {
            // List<string> locs = CreateCircle2(centerLat, centerLon, radius);
            List<string> locs = CreateCircle(centerLat, centerLon, radius);
            return string.Format("Polygon(({0}))", string.Join(",", locs));
        }

        private static List<string> CreateCircle2(double centerLat, double centerLon, double radius)
        {
            List<string> locs = new List<string>();
            //var lat1 = lat * Math.PI / 180.0; // From Degrees To Radians
            //var lon1 = lon * Math.PI / 180.0;

            // From Degrees To Radians
            var lat1 = ToRadian(centerLat);
            var lon1 = ToRadian(centerLon);

            var d = radius / 3956;
            for (int x = 0; x <= 360; x++)
            {
                // Calculate Latitude.
                var tc = (x / 90) * Math.PI / 2;
                var latC = Math.Asin(Math.Sin(lat1) * Math.Cos(d) + Math.Cos(lat1) * Math.Sin(d) * Math.Cos(tc));
                latC = ToDegrees(latC); // From radians To degrees

                // Calculate longitude.
                double lonC;
                if (Math.Cos(lat1) == 0)
                {
                    lonC = centerLon; // endpoint a pole 
                }
                else
                {
                    lonC = ((lon1 - Math.Asin(Math.Sin(tc) * Math.Sin(d) / Math.Cos(lat1)) + Math.PI) % (2 * Math.PI)) - Math.PI;
                }

                lonC = ToDegrees(lonC);
                locs.Add(string.Format("{0} {1}", lonC, latC));
            }
            return locs;
        }

        public static List<string> CreateCircle(double centerLat, double centerLon, double radius)
        {
            var earthRadius = 3956.0;
            var lat = ToRadian(centerLat); //radians            
            var lng = ToRadian(centerLon); //radians           
            var d = radius / earthRadius; // d = angular distance covered on earth's surface            
            var locations = new List<string>();
            for (var x = 0; x <= 360; x++)
            {
                var brng = ToRadian(x);
                var latRadians = Math.Asin(Math.Sin(lat) * Math.Cos(d) + Math.Cos(lat) * Math.Sin(d) * Math.Cos(brng));
                var lngRadians = lng + Math.Atan2(Math.Sin(brng) * Math.Sin(d) * Math.Cos(lat), Math.Cos(d) - Math.Sin(lat) * Math.Sin(latRadians));

                locations.Add(string.Format("{0} {1}", ToDegrees(lngRadians), ToDegrees(latRadians)));
            }

            return locations;
        }

        public static double ToRadian(double degrees)
        {
            return degrees * (Math.PI / 180);
        }

        public static double ToDegrees(double radians)
        {
            return radians * (180 / Math.PI);
        }
    }
}
