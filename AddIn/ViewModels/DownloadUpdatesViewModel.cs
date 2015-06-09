//-----------------------------------------------------------------------
// <copyright file="DownloadUpdatesViewModel.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using Microsoft.Research.Wwt.Excel.Addin.Properties;
namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// View Model for the DownloadUpdates Button on the custom task pane
    /// </summary>
    public class DownloadUpdatesViewModel : PropertyChangeBase
    {
        #region Private Properties
        private bool isDownloadUpdatesEnabled;
        private bool isDownloadUpdatesVisible;
        private string downloadUpdatesLabel;
        #endregion 

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the DownloadUpdatesViewModel class
        /// </summary>
        public DownloadUpdatesViewModel()
        {
            this.isDownloadUpdatesEnabled = false;
            this.IsDownloadUpdatesVisible = false;
            this.DownloadUpdatesLabel = Resources.DownloadUpdatesButtonLabel;
        }

        #endregion

        #region Properties
        /// <summary>
        /// Gets or sets a value indicating whether the download updates button is enabled
        /// </summary>
        public bool IsDownloadUpdatesEnabled
        {
            get
            {
                return this.isDownloadUpdatesEnabled;
            }
            set
            {
                this.isDownloadUpdatesEnabled = value;
                OnPropertyChanged("IsDownloadUpdatesEnabled");
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the download updates button is visible
        /// </summary>
        public bool IsDownloadUpdatesVisible
        {
            get
            {
                return this.isDownloadUpdatesVisible;
            }
            set
            {
                this.isDownloadUpdatesVisible = value;
                OnPropertyChanged("IsDownloadUpdatesVisible");
            }
        }

        /// <summary>
        /// Gets or sets the text that appears on the download updates button
        /// </summary>
        public string DownloadUpdatesLabel
        {
            get
            {
                return this.downloadUpdatesLabel;
            }
            set
            {
                this.downloadUpdatesLabel = value;
                OnPropertyChanged("DownloadUpdatesLabel");
            }
        }
        #endregion
    }
}
