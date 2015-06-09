// -----------------------------------------------------------------------
// <copyright file="FetchClimateOutputModel.cs" company="AditiTechnologies Pvt Ltd">
// Fetch climate output model class
// </copyright>
// -----------------------------------------------------------------------

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Fetch climate output model class.
    /// </summary>
    public class FetchClimateOutputModel
    {
        #region Public properties 
        /// <summary>
        /// Gets or sets minimum Latitude  value.
        /// </summary>
        public double MinLatitude { get; set; }

        /// <summary>
        /// Gets or sets minimum Longitude  value.
        /// </summary>
        public double MinLongitude { get; set; }

        /// <summary>
        /// Gets or sets maximum Latitude  value.
        /// </summary>
        public double MaxLatitude { get; set; }

        /// <summary>
        /// Gets or sets maximum Longitude  value.
        /// </summary>
        public double MaxLongitude { get; set; }

        /// <summary>
        /// Gets or sets maximum Precipitation  value.
        /// </summary>
        public double Precipitation { get; set; }

        /// <summary>
        /// Gets or sets maximum Temperature value.
        /// </summary>
        public double Temperature { get; set; }

        /// <summary>
        /// Gets Altitude value.
        /// </summary>
        public double Altitude
        {
            get
            {
                if (!double.IsNaN(Precipitation))
                {
                    return Precipitation * 1000;
                }
                else
                {
                    return 0;
                }
            }
        }

        /// <summary>
        /// Gets Geometry value.
        /// </summary>
        public string Geometry
        {
            get
            {
                return "Polygon((" + MinLongitude + " " + MinLatitude + "," + MinLongitude + " " + MaxLatitude + "," + MaxLongitude + " " + MaxLatitude + "," + MaxLongitude + " " + MinLatitude + "," + MinLongitude + " " + MinLatitude + "))";
            }
        }

        /// <summary>
        /// Gets Color value.
        /// </summary>
        public string Color
        {
            get
            {
                if (double.IsNaN(Temperature))
                {
                    return "Transparent";
                }
                else if (Temperature < 0)
                {
                    return "50%White";
                }
                else if (Temperature < 3)
                {
                    return "50%Blue";
                }
                else if (Temperature < 6)
                {
                    return "50%Cyan";
                }
                else if (Temperature < 9)
                {
                    return "50%Green";
                }
                else if (Temperature < 12)
                {
                    return "50%Yellow";
                }
                else if (Temperature < 15)
                {
                    return "50%Orange";
                }
                else
                {
                    return "50%Red";
                }
            }
        }

        #endregion

        #region Constructor 

        /// <summary>
        /// Initializes a new instance of the FetchClimateOutputModel class.
        /// </summary>
        /// <param name="latMin">Minimum Latitude.</param>
        /// <param name="latMax">Maximum Latitude.</param>
        /// <param name="longMin">Minimum Longitude.</param>
        /// <param name="longMax">Maximum Longitude.</param>
        /// <param name="precipitation">Precipitation value.</param>
        /// <param name="temp">Temperature value.</param>
        public FetchClimateOutputModel(double latMin, double latMax, double longMin, double longMax, double precipitation, double temp)
        {
            this.MinLatitude = latMin;
            this.MinLongitude = longMin;
            this.MaxLatitude = latMax;
            this.MaxLongitude = longMax;
            this.Precipitation = precipitation;
            this.Temperature = temp;
        }

        #endregion
    }
}
