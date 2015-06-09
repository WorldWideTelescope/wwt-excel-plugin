//-----------------------------------------------------------------------
// <copyright file="EventState.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;

namespace Microsoft.Research.Wwt.Excel.Common.Tests
{
    /// <summary>
    /// Helper class used by the EventHelperTest class
    /// </summary>
    public class EventState : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the EventState class
        /// </summary>
        public EventState()
        {
            this.EventFired = false;
        }

        /// <summary>
        /// Gets or sets a value indicating whether the event has been fired
        /// </summary>
        public bool EventFired { get; set; }
    }
}
