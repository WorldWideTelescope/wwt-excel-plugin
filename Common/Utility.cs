//-----------------------------------------------------------------------
// <copyright file="Utility.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Security;
using Microsoft.Research.Wwt.Excel.Common.Properties;
using Microsoft.Win32;

namespace Microsoft.Research.Wwt.Excel.Common
{
    public static class Utility
    {
        private static object syncRoot = new object();

        /// <summary>
        /// Bring WWT in focus
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2122:DoNotIndirectlyExposeMethodsWithLinkDemands", Justification = "Process class needs link demand.")]
        public static void ShowWWT()
        {
            try
            {
                lock (syncRoot)
                {
                    // Set the exe name default value to start with
                    string processName = "WWTExplorer";
                    if (!string.IsNullOrWhiteSpace(Globals.WWTApplicationPath))
                    {
                        // If value from registry
                        processName = Path.GetFileNameWithoutExtension(Globals.WWTApplicationPath);
                    }

                    Process[] previousProcess = Process.GetProcessesByName(processName);
                    foreach (Process proc in previousProcess)
                    {
                        if (proc != null && !proc.HasExited)
                        {
                            // Bring it in focus
                            NativeMethods.SetForegroundWindow(proc.MainWindowHandle);
                            NativeMethods.ShowWindow(proc.MainWindowHandle, 3);
                            break;
                        }
                    }
                }
            }
            catch (ArgumentException ex)
            {
                Logger.LogException(ex);
            }
            catch (InvalidOperationException ex)
            {
                Logger.LogException(ex);
            }
            catch (Win32Exception ex)
            {
                Logger.LogException(ex);
            }
            catch (NotSupportedException ex)
            {
                Logger.LogException(ex);
            }
        }

        /// <summary>
        /// Get the path of the WWT application from the registry
        /// using the invariant .wtml extension
        /// </summary>
        /// <returns>string containing WWT Path</returns>
        public static void SetWWTApplicationPath()
        {
            Common.Globals.WWTApplicationPath = string.Empty;

            // read the path of the executable from the registry
            try
            {
                var regKeyExtension = Registry.ClassesRoot.OpenSubKey(Constants.WtmlFileExtension);
                if (regKeyExtension != null)
                {
                    var classInfo = regKeyExtension.GetValue(null);
                    if (classInfo != null)
                    {
                        var regKeyWWTPath = Registry.ClassesRoot.OpenSubKey((string)classInfo + Constants.ShellOpenCommandRegistryKey);
                        if (regKeyWWTPath != null)
                        {
                            var registryPath = regKeyWWTPath.GetValue(null);
                            if (registryPath != null)
                            {
                                string[] registryPathParts = registryPath.ToString().Split(new char[] { '\"' });
                                if (registryPathParts.Length > 1)
                                {
                                    Common.Globals.WWTApplicationPath = registryPathParts[1];
                                }
                            }
                        }
                    }
                }
            }
            catch (ArgumentNullException ex)
            {
                Logger.LogException(ex);
            }
            catch (ObjectDisposedException ex)
            {
                Logger.LogException(ex);
            }
            catch (SecurityException ex)
            {
                Logger.LogException(ex);
            }
            catch (UnauthorizedAccessException ex)
            {
                Logger.LogException(ex);
            }
            catch (IOException ex)
            {
                Logger.LogException(ex);
            }
        }

        /// <summary>
        /// Check if WWT is installed on the local machine
        /// </summary>
        /// <returns>True if the registry contains an entry for .wtml in classes root</returns>
        public static void IsWWTInstalled()
        {
            try
            {
                if (TargetMachine.DefaultIP.ToString() == Common.Globals.TargetMachine.MachineIP.ToString())
                {
                    var regKeyExtension = Registry.ClassesRoot.OpenSubKey(Constants.WtmlFileExtension);
                    if (regKeyExtension == null)
                    {
                        throw new CustomException(Properties.Resources.WWTNotInstalledError, true, ErrorCodes.Code100004);
                    }
                }
            }
            catch (ArgumentNullException ex)
            {
                Logger.LogException(ex);
                throw new CustomException(String.Format(CultureInfo.InvariantCulture, Resources.ErrorReadingRegistry, Properties.Resources.ProductNameShort), true);
            }
            catch (ObjectDisposedException ex)
            {
                Logger.LogException(ex);
                throw new CustomException(String.Format(CultureInfo.InvariantCulture, Resources.ErrorReadingRegistry, Properties.Resources.ProductNameShort), true);
            }
            catch (SecurityException ex)
            {
                Logger.LogException(ex);
                throw new CustomException(String.Format(CultureInfo.InvariantCulture, Resources.ErrorReadingRegistry, Properties.Resources.ProductNameShort), true);
            }
        }
    }
}
