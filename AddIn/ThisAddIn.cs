//-----------------------------------------------------------------------
// <copyright file="ThisAddIn.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using System.Diagnostics.CodeAnalysis;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// This class is responsible for loading the plug-in and listening to events.
    /// </summary>
    [ExcludeFromCodeCoverage]
    public partial class ThisAddIn
    {
        /// <summary>
        /// Gets or sets the excel application object.
        /// </summary>
        internal static Microsoft.Office.Interop.Excel.Application ExcelApplication { get; set; }

        /// <summary>
        /// Add-in startup event
        /// </summary>
        /// <param name="sender">
        /// Event sender
        /// </param>
        /// <param name="e">
        /// object that contains the data that can be passed from the event sender.
        /// Inherited from the System.EventArgs class
        /// </param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        /// <summary>
        /// This function is triggered when the plug-in gets released. 
        /// The shutdown procedure is handled by VSTO. Use this function if you need to perform additional 
        /// actions before releasing plug-in and closing.
        /// </summary>
        /// <param name="sender">
        /// Event sender
        /// </param>
        /// <param name="e">
        /// object that contains the data that can be passed from the event sender.
        /// Inherited from the System.EventArgs class
        /// </param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
