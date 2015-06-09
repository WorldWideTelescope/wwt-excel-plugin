// -----------------------------------------------------------------------
// <copyright file="FetchClimateInputModel.cs" company="AditiTechnologies Pvt Ltd">
// Model class for fetch climate input values.
// </copyright>
// -----------------------------------------------------------------------

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    ///  Model class for fetch climate input values.
    /// </summary>
    public class FetchClimateInputModel
    {
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
        /// Gets or sets delta Latitude  value.
        /// </summary>
        public double DeltaLatitude { get; set; }

        /// <summary>
        /// Gets or sets delta Longitude  value.
        /// </summary>
        public double DeltaLongitude { get; set; }

    }
}
