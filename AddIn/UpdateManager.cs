//-----------------------------------------------------------------------
// <copyright file="UpdateManager.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Microsoft.Research.Wwt.Excel.Addin.Properties;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Handles the auto update feature of the addin
    /// </summary>
    internal class UpdateManager : IDisposable
    {
        /// <summary>
        /// The link from which to download the updated WWT Excel Add-In.
        /// </summary>
        private string downloadLink;

        /// <summary>
        /// The local path of the downloaded Excel Add-In installer.
        /// </summary>
        private string downloadedUpdateFile;

        private BackgroundWorker checkUpdatesWorker;
        private BackgroundWorker downloadLinkWorker;
        private BackgroundWorker installUpdateWorker;

        /// <summary>
        /// Track whether Dispose has been called.
        /// </summary>
        private bool disposed = false;

        /// <summary>
        /// Initializes a new instance of the UpdateManager class
        /// </summary>
        internal UpdateManager()
        {
            this.checkUpdatesWorker = new BackgroundWorker();
            this.downloadLinkWorker = new BackgroundWorker();
            this.installUpdateWorker = new BackgroundWorker();

            this.checkUpdatesWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.OnCheckUpdatesWorkerDoWork);
            this.checkUpdatesWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.OnCheckUpdatesWorkerRunWorkerCompleted);
            this.downloadLinkWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.OnDownloadLinkWorkerDoWork);
            this.downloadLinkWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.OnDownloadLinkWorkerRunWorkerCompleted);
            this.installUpdateWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.OnInstallUpdateWorkerDoWork);
            this.installUpdateWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.OnInstallUpdateWorkerRunWorkerCompleted);
        }

        internal event EventHandler UpdateAvailable;
        internal event EventHandler DownloadCompleted;
        internal event EventHandler InstallationCompleted;

        /// <summary>
        /// Start the check for updates worker
        /// </summary>
        public void CheckForUpdates()
        {
            try
            {
                if (this.checkUpdatesWorker != null)
                {
                    this.checkUpdatesWorker.RunWorkerAsync();
                }
                else
                {
                    throw new CustomException(Resources.DefaultErrorMessage, false);
                }
            }
            catch (InvalidOperationException)
            {
                throw new CustomException(Resources.DefaultErrorMessage, false);
            }
        }

        /// <summary>
        /// Start the download link worker
        /// </summary>
        public void DownloadUpdates()
        {
            try
            {
                if (this.downloadLinkWorker != null)
                {
                    this.downloadLinkWorker.RunWorkerAsync();
                }
                else
                {
                    throw new CustomException(Resources.DefaultErrorMessage, false);
                }
            }
            catch (InvalidOperationException)
            {
                throw new CustomException(Resources.DefaultErrorMessage, false);
            }
        }

        /// <summary>
        /// Start the install updates worker
        /// </summary>
        public void InstallUpdates()
        {
            try
            {
                if (this.installUpdateWorker != null)
                {
                    this.installUpdateWorker.RunWorkerAsync();
                }
                else
                {
                    throw new CustomException(Resources.DefaultErrorMessage, false);
                }
            }
            catch (InvalidOperationException)
            {
                throw new CustomException(Resources.DefaultErrorMessage, false);
            }
        }

        /// <summary>
        /// Part of IDisposable Interface
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Part of IDisposable Interface
        /// </summary>
        /// <param name="disposing">True if called from code</param>
        protected virtual void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!this.disposed)
            {
                // If disposing equals true, dispose all managed
                // and unmanaged resources.
                if (disposing)
                {
                    this.checkUpdatesWorker.Dispose();
                    this.downloadLinkWorker.Dispose();
                    this.installUpdateWorker.Dispose();
                }

                // Note disposing has been done.
                disposed = true;
            }
        }

        /// <summary>
        /// DoWork event handler for the check updates background worker
        /// This is run asynchronously
        /// </summary>
        /// <param name="sender">Check updates worker</param>
        /// <param name="e">Do work event args</param>
        private void OnCheckUpdatesWorkerDoWork(object sender, DoWorkEventArgs e)
        {
            string downloadUrl = WWTManager.CheckForUpdates();
            if (!string.IsNullOrEmpty(downloadUrl))
            {
                e.Result = downloadUrl;
            }
        }

        /// <summary>
        /// RunWorkerCompleted Event handler for the check updates Background Worker
        /// This is called after the asynchronous call DoWork is completed
        /// </summary>
        /// <param name="sender">check updates worker</param>
        /// <param name="e">Run worker completed event args</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        private void OnCheckUpdatesWorkerRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                this.downloadLink = (string)e.Result;
                if (!string.IsNullOrEmpty(downloadLink))
                {
                    this.UpdateAvailable.OnFire(this, new EventArgs());
                }
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
            }
        }

        /// <summary>
        /// RunWorkerCompleted Event handler for the download link Background Worker
        /// This is called after the asynchronous call DoWork is completed
        /// </summary>
        /// <param name="sender">Download link worker</param>
        /// <param name="e">Run worker completed event args</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        private void OnDownloadLinkWorkerRunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            try
            {
                this.downloadedUpdateFile = (string)e.Result;
                if (!string.IsNullOrEmpty(downloadedUpdateFile))
                {
                    this.DownloadCompleted.OnFire(this, new EventArgs());
                }
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
            }
        }

        /// <summary>
        /// DoWork event handler for the download link background worker
        /// This is run asynchronously
        /// </summary>
        /// <param name="sender">Download link worker</param>
        /// <param name="e">Do work event args</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        private void OnDownloadLinkWorkerDoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            try
            {
                Uri downloadUri = new Uri(this.downloadLink);

                // Ensure that the temp directory exists
                if (!Directory.Exists(Path.GetTempPath()))
                {
                    Directory.CreateDirectory(Path.GetTempPath());
                }

                // Gets the filename from Uri and appends it with temp path.
                string localFilePath = Path.Combine(Path.GetTempPath(), downloadUri.Segments[downloadUri.Segments.Length - 1]);

                if (WWTManager.DownloadFile(downloadUri, localFilePath))
                {
                    e.Result = string.Format(CultureInfo.InvariantCulture, localFilePath, Path.GetTempPath());
                }
            }
            catch (ArgumentNullException ex)
            {
                Logger.LogException(ex);
            }
            catch (FormatException ex)
            {
                Logger.LogException(ex);
            }
            catch (System.Security.SecurityException ex)
            {
                Logger.LogException(ex);
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
            }
        }

        /// <summary>
        /// DoWork Event handler for the Install Update Background Worker
        /// This is called asynchronously
        /// </summary>
        /// <param name="sender">Install Update Background Worker</param>
        /// <param name="e">Do work event args</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        private void OnInstallUpdateWorkerDoWork(object sender, DoWorkEventArgs e)
        {
            using (Process msiexec = new Process())
            {
                msiexec.StartInfo.Arguments = "/i " + this.downloadedUpdateFile;
                msiexec.StartInfo.FileName = "msiexec";
                try
                {
                    msiexec.Start();

                    // This will make sure user exit installation process before enabling the update button again.
                    msiexec.WaitForExit();
                }
                catch (InvalidOperationException ex)
                {
                    Logger.LogException(ex);
                }
                catch (System.ComponentModel.Win32Exception ex)
                {
                    Logger.LogException(ex);
                }
                catch (Exception exception)
                {
                    Logger.LogException(exception);
                }
            }
        }

        /// <summary>
        /// Run Worker Completed Event Handler for the Install Update Background Worker
        /// </summary>
        /// <param name="sender">Install Update Background Worker</param>
        /// <param name="e">Run Worker Completed Event Arguments</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General exception needs to be caught for logging in case of any rare scenario.")]
        private void OnInstallUpdateWorkerRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                this.InstallationCompleted.OnFire(this, new EventArgs());
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
            }
        }
    }
}
