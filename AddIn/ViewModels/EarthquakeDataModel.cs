// -----------------------------------------------------------------------
// <copyright file="EarthquakeDataModel.cs">
//  View model for fetch climate view.
// </copyright>
// -----------------------------------------------------------------------

namespace Microsoft.Research.Wwt.Excel.Addin
{
    using System;
    using System.Linq;
    using System.Collections.Generic;
    using System.Windows.Input;
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Research.Wwt.Excel.Common;
    using Microsoft.Research.Wwt.Excel.Addin.Properties;
    using System.Collections.ObjectModel;
    using System.ComponentModel;

    public enum PriorityType
    {
        time,
        size
    }

    public class EarthquakeDataModel : PropertyChangeBase
    {
        private string magnitudeMin;
        public string MagnitudeMin
        {
            get { return magnitudeMin; }
            set
            {
                magnitudeMin = value;
                OnPropertyChanged("MagnitudeMin");
            }
        }

        private string magnitudeMax;
        public string MagnitudeMax
        {
            get { return magnitudeMax; }
            set
            {
                magnitudeMax = value;
                OnPropertyChanged("MagnitudeMax");
            }
        }

        private string depthMin;
        public string DepthMin
        {
            get { return depthMin; }
            set
            {
                depthMin = value;
                OnPropertyChanged("DepthMin");
            }
        }

        private string depthMax;
        public string DepthMax
        {
            get { return depthMax; }
            set
            {
                depthMax = value;
                OnPropertyChanged("DepthMax");
            }
        }

        private string startDate;
        public string StartDate
        {
            get { return startDate; }
            set
            {
                startDate = value;
                OnPropertyChanged("StartDate");
            }
        }

        private string endDate;
        public string EndDate
        {
            get { return endDate; }
            set
            {
                endDate = value;
                OnPropertyChanged("EndDate");
            }
        }

        public EarthquakeDataModel()
        {
            this.priority = PopulatePriorityType();
            this.selectedPriority = this.priority[0];

            this.displayCount = PopulateDisplayCount();
            this.selectedDisplayCount = this.displayCount[0];

            this.magnitudeMax = "10";
            this.magnitudeMin = "1";
            this.depthMin = "1";
            this.depthMax = "900";
            this.startDate = "2009/01/01";
            this.EndDate = DateTime.Now.ToString(@"yyyy/MM/dd");


        }

        private Collection<KeyValuePair<PriorityType, string>> priority;

        public ReadOnlyCollection<KeyValuePair<PriorityType, string>> Priority
        {
            get
            {
                return new ReadOnlyCollection<KeyValuePair<PriorityType, string>>(this.priority);
            }
        }

        private Collection<KeyValuePair<int, string>> displayCount;

        public ReadOnlyCollection<KeyValuePair<int, string>> DisplayCount
        {
            get
            {
                return new ReadOnlyCollection<KeyValuePair<int, string>>(this.displayCount);
            }
        }

        private KeyValuePair<int, string> selectedDisplayCount;
        /// <summary>
        /// Gets or sets SelectedDistanceUnit
        /// </summary>
        public KeyValuePair<int, string> SelectedDisplayCount
        {
            get
            {
                return this.selectedDisplayCount;
            }
            set
            {
                selectedDisplayCount = value;
                OnPropertyChanged("SelectedDisplayCount");
            }
        }

        private KeyValuePair<PriorityType, string> selectedPriority;
        /// <summary>
        /// Gets or sets SelectedDistanceUnit
        /// </summary>
        public KeyValuePair<PriorityType, string> SelectedPriority
        {
            get
            {
                return this.selectedPriority;
            }
            set
            {
                selectedPriority = value;
                OnPropertyChanged("SelectedPriority");
            }
        }

        private static Collection<KeyValuePair<PriorityType, string>> PopulatePriorityType()
        {
            Collection<KeyValuePair<PriorityType, string>> priorities = new Collection<KeyValuePair<PriorityType, string>>();
            priorities.Add(new KeyValuePair<PriorityType, string>(PriorityType.time, "Newer Events"));
            priorities.Add(new KeyValuePair<PriorityType, string>(PriorityType.size, "Larger Events"));
            return priorities;
        }

        private static Collection<KeyValuePair<int, string>> PopulateDisplayCount()
        {
            Collection<KeyValuePair<int, string>> displayCount = new Collection<KeyValuePair<int, string>>();
            displayCount.Add(new KeyValuePair<int, string>(100, "100 Events"));
            displayCount.Add(new KeyValuePair<int, string>(200, "200 Events"));
            displayCount.Add(new KeyValuePair<int, string>(300, "300 Events"));
            displayCount.Add(new KeyValuePair<int, string>(400, "400 Events"));
            displayCount.Add(new KeyValuePair<int, string>(500, "500 Events"));
            displayCount.Add(new KeyValuePair<int, string>(600, "600 Events"));
            displayCount.Add(new KeyValuePair<int, string>(700, "700 Events"));
            displayCount.Add(new KeyValuePair<int, string>(800, "800 Events"));
            displayCount.Add(new KeyValuePair<int, string>(900, "900 Events"));
            displayCount.Add(new KeyValuePair<int, string>(1000, "1000 Events"));
            displayCount.Add(new KeyValuePair<int, string>(1250, "1250 Events"));
            displayCount.Add(new KeyValuePair<int, string>(1500, "1500 Events"));
            displayCount.Add(new KeyValuePair<int, string>(1750, "1750 Events"));
            displayCount.Add(new KeyValuePair<int, string>(2000, "2000 Events"));
            displayCount.Add(new KeyValuePair<int, string>(2250, "2250 Events"));
            displayCount.Add(new KeyValuePair<int, string>(2500, "2500 Events"));
            displayCount.Add(new KeyValuePair<int, string>(2750, "2750 Events"));
            displayCount.Add(new KeyValuePair<int, string>(3000, "3000 Events"));
            displayCount.Add(new KeyValuePair<int, string>(3500, "3500 Events"));
            displayCount.Add(new KeyValuePair<int, string>(4000, "4000 Events"));
            displayCount.Add(new KeyValuePair<int, string>(4500, "4500 Events"));
            displayCount.Add(new KeyValuePair<int, string>(5000, "5000 Events"));
            displayCount.Add(new KeyValuePair<int, string>(10000, "10000 Events"));
            displayCount.Add(new KeyValuePair<int, string>(20000, "20000 Events"));
            displayCount.Add(new KeyValuePair<int, string>(3000, "30000 Events"));
            displayCount.Add(new KeyValuePair<int, string>(40000, "40000 Events"));
            displayCount.Add(new KeyValuePair<int, string>(50000, "50000 Events"));
            displayCount.Add(new KeyValuePair<int, string>(75000, "75000 Events"));
            displayCount.Add(new KeyValuePair<int, string>(100000, "100000 Events"));
            return displayCount;
        }
    }
}
