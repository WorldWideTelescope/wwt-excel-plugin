// -----------------------------------------------------------------------
// <copyright file="FetchClimateAPIUtility.cs" company="AditiTechnologies Pvt Ltd">
// Utility class to get the data from Fetch climate API.
// </copyright>
// -----------------------------------------------------------------------

namespace Microsoft.Research.Wwt.Excel.Addin
{
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Research.Science.Data;
    using Microsoft.Research.Wwt.Excel.Common;

    /// <summary>
    /// Utility class to get the data from Fetch climate API.
    /// </summary>
    public static class FetchClimateAPIUtility
    {
        /// <summary>
        /// Method to get the list of precipitation and temparature values.
        /// </summary>
        /// <param name="latMin">Min latitude.</param>
        /// <param name="latMax">Max latitude.</param>
        /// <param name="longMin">Min longitude.</param>
        /// <param name="longMax">Max longitude.</param>
        /// <param name="dlat">Delta latitude.</param>
        /// <param name="dlong">Delta longitude.</param>
        /// <returns>List of FetchClimateOutputModel objects.</returns>
        public static List<FetchClimateOutputModel> GetPrecipitationAndTemp(double latMin, double latMax, double longMin, double longMax, double dlat, double dlong)
        {
            List<FetchClimateOutputModel> lstFetchClimateValues = new List<FetchClimateOutputModel>();
            double longMinTemp = longMin;

            try
            {
                // Creating collection of latitude and longitude values dependingon the user inputs.
                while (latMin < latMax)
                {
                    while (longMin < longMax)
                    {
                        lstFetchClimateValues.Add(new FetchClimateOutputModel(latMin, latMin + dlat, longMin, longMin + dlong, 0.00, 0.00));
                        longMin += dlong;
                    }

                    latMin += dlat;
                    longMin = longMinTemp;
                }

                // Getting list of precipitation values from fetch climate API
                double[] precipitation = ClimateService.FetchClimate(ClimateParameter.FC_PRECIPITATION, lstFetchClimateValues.Select(o => o.MinLatitude).ToArray(), lstFetchClimateValues.Select(o => o.MaxLatitude).ToArray(), lstFetchClimateValues.Select(o => o.MinLongitude).ToArray(), lstFetchClimateValues.Select(o => o.MaxLongitude).ToArray());

                // Getting list of temparature values from fetch climate API
                double[] temp = ClimateService.FetchClimate(ClimateParameter.FC_TEMPERATURE, lstFetchClimateValues.Select(o => o.MinLatitude).ToArray(), lstFetchClimateValues.Select(o => o.MaxLatitude).ToArray(), lstFetchClimateValues.Select(o => o.MinLongitude).ToArray(), lstFetchClimateValues.Select(o => o.MaxLongitude).ToArray());

                // Assigning precipitation and temparature values.
                lstFetchClimateValues.ForEach(location =>
                {
                    int index = lstFetchClimateValues.IndexOf(location);
                    location.Precipitation = precipitation[index];
                    location.Temperature = temp[index];
                });
            }
            catch
            {
                lstFetchClimateValues = null;
            }

            return lstFetchClimateValues;
        }
    }
}
