//-----------------------------------------------------------------------
// <copyright file="Globals.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Diagnostics;

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Global state to be used by all
    /// </summary>
    public static class Globals
    {
        /// <summary>
        /// addinTraceSource instance
        /// </summary>
        private static readonly TraceSource addinTraceSource = new TraceSource("WWTEarthAddin");

        /// <summary>
        /// WWTManager instance. Default instance created with WWTRequest object.
        /// </summary>
        private static WWTManager wwtManager = new WWTManager(new WWTRequest());

        /// <summary>
        /// Gets the WWTManager instance.
        /// </summary>
        public static WWTManager WWTManager
        {
            get { return wwtManager; }
        }

        /// <summary>
        /// Gets or sets the TargetMachine.
        /// </summary>
        public static TargetMachine TargetMachine { get; set; }

        /// <summary>
        /// Gets or sets the installed location of the WWT Application
        /// </summary>
        public static string WWTApplicationPath { get; set; }

        /// <summary>
        /// Gets trace source which will be used for tracing and logging
        /// </summary>
        public static TraceSource AddinTraceSource
        {
            get { return addinTraceSource; }
        }
    }
}
