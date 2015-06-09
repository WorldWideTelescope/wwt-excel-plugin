//-----------------------------------------------------------------------
// <copyright file="NativeMethods.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Runtime.InteropServices;

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Definition of Win32 API calls
    /// </summary>
    internal static class NativeMethods
    {
        [DllImport("User32.dll")]
        internal static extern IntPtr SetForegroundWindow(IntPtr windowHandle);

        [DllImport("User32.dll")]
        internal static extern IntPtr ShowWindow(IntPtr windowHandle, int nCmdShow);
    }
}
