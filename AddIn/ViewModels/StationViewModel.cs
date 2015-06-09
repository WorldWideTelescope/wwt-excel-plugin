
namespace Microsoft.Research.Wwt.Excel.Addin
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Windows.Input;
    using Microsoft.Research.Wwt.Excel.Common;

    public class StationViewModel : PropertyChangeBase
    {
        public StationDataModel StationData { get; set; }

        public EarthquakeDataModel EarthquakeData { get; set; } 

        private ICommand saveCommand;

        private ICommand getEarthquakeDataCommand;

        public ICommand SaveCommand
        {
            get { return this.saveCommand; }
        }

        public ICommand GetEarthquakeDataCommand
        {
            get { return this.getEarthquakeDataCommand; }
        }

        public StationViewModel() 
        {
            StationData = new StationDataModel();
            EarthquakeData = new EarthquakeDataModel(); 
            this.saveCommand = new SaveCommandHandler(this);
            this.getEarthquakeDataCommand = new GetEarthquakeDataCommandHandler(this);
        }

        #region CustomEvent

        /// <summary>
        /// Capture View window close request
        /// </summary>
        public event EventHandler RequestClose;

        #endregion

        #region Public methods

        /// <summary>
        /// Raise window close event
        /// </summary>
        public void OnRequestClose()
        {
            if (RequestClose != null)
            {
                RequestClose(this, EventArgs.Empty);
            }
        }

        #endregion

        private class SaveCommandHandler : RelayCommand
        {
            private StationViewModel parent;

            public SaveCommandHandler(StationViewModel stationViewModel)
            {
                this.parent = stationViewModel;
            }

            public override void Execute(object parameter)
            {
                if (this.parent != null)
                {
                    WorkflowController.Instance.GetStationData(this.parent);
                }
            }
        }

        private class GetEarthquakeDataCommandHandler : RelayCommand
        {
            private StationViewModel parent;

            public GetEarthquakeDataCommandHandler(StationViewModel stationViewModel)
            {
                this.parent = stationViewModel;
            }

            public override void Execute(object parameter)
            {
                if (this.parent != null)
                {
                    WorkflowController.Instance.GetEarthquakeData(this.parent);
                }
            }
        }
    }
}
