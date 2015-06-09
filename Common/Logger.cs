//-----------------------------------------------------------------------
// <copyright file="Logger.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;

namespace Microsoft.Research.Wwt.Excel.Common
{
    public static class Logger
    {
        /// <summary>
        /// Constructs and logs the exception message.
        /// </summary>
        /// <param name="exception">The Exception object.</param>
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "Consume any exception while logging.")]
        public static void LogException(Exception exception)
        {
            if (exception != null)
            {
                try
                {
                    string traceMessage = DateTime.Now + " : " + exception.Message;
                    if (exception.InnerException != null)
                    {
                        traceMessage += " : " + exception.InnerException.Message;
                    }

                    Globals.AddinTraceSource.TraceEvent(TraceEventType.Error, exception.GetHashCode(), traceMessage);
                }
                catch (Exception)
                {
                    // Consume any exception while logging as it cannot be logged any more.
                }
            }
        }
    }
}
