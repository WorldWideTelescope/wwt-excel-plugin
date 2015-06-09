
namespace Microsoft.Research.Wwt.Excel.Addin
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    public class StationDataModel : PropertyChangeBase
    {
        public StationDataModel()
        {
            // Set Default
            this.network = "IU";
            this.station = "ANMO";
            this.location = "00";
            this.channel = "LH?,BH*";
            this.startDate = "1997-06-07";
            this.endDate = "2011-06-07";
            this.level = "sta";
        }

        private string network;
        public string Network
        {
            get { return network; }
            set
            {
                network = value;
                OnPropertyChanged("Network");
            }
        }

        private string station;
        public string Station
        {
            get { return station; }
            set
            {
                station = value;
                OnPropertyChanged("Station");
            }
        }

        private string location;
        public string Location
        {
            get { return location; }
            set
            {
                location = value;
                OnPropertyChanged("Location");
            }
        }

        private string channel;
        public string Channel
        {
            get { return channel; }
            set
            {
                channel = value;
                OnPropertyChanged("Channel");
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

        private string level;
        public string Level
        {
            get { return level; }
            set
            {
                level = value;
                OnPropertyChanged("Level");
            }
        }
    }
}
