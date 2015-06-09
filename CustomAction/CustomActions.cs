//-----------------------------------------------------------------------
// <copyright file="CustomActions.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Win32;

namespace Microsoft.Research.Wwt.Excel.Installer.CustomAction
{
    /// <summary>
    /// Enum to indicate the Processor Architecture.
    /// </summary>
    public enum ProcessorType
    {
        /// <summary>
        /// 32 bit processor.
        /// </summary>
        x86,

        /// <summary>
        /// 64 bit processor.
        /// </summary>
        x64
    }

    /// <summary>
    /// Custom actions that aids install.
    /// </summary>
    public static class CustomActions
    {

        /// <summary>
        /// Holds the Key for custom data (office version).
        /// </summary>
        private const string OfficeVersion = "OfficeVersion";

        /// <summary>
        /// Office 2007 version number.
        /// </summary>
        private const string Office2007Version = "12.0";

        /// <summary>
        /// Office 2010 version number.
        /// </summary>
        private const string Office2010Version = "15.0";

        /// <summary>
        /// Property to check if Excel is running.
        /// </summary>
        private const string ExcelRunning = "IsExcelRunning";

        /// <summary>
        /// The registry key of WWT application
        /// </summary>
        private const string WWTRegPath = @"Software\Microsoft\WWT";

        /// <summary>
        /// The registry key of WWT application in x64
        /// </summary>
        private const string WWT64RegPath = @"Software\Wow6432Node\Microsoft\WWT";

        /// <summary>
        /// Excel 2007 registry key
        /// </summary>
        private const string Excel2007RegKey = @"Software\Microsoft\Office\12.0\Excel\InstallRoot";

        /// <summary>
        /// Excel 2010 registry key 
        /// </summary>
        private const string Excel2010RegKey = @"Software\Microsoft\Office\14.0\Excel\InstallRoot";

        /// <summary>
        /// Excel 2013 registry key 
        /// </summary>
        private const string Excel2013RegKey = @"Software\Microsoft\Office\15.0\Excel\InstallRoot";

        /// <summary>
        /// VSTO 2010 Registry Path
        /// </summary>
        private const string VSTO2010RegKey = @"Software\Microsoft\VSTO Runtime Setup\v4R";

        /// <summary>
        /// Dot net framework 4.0 full registry path
        /// </summary>
        private const string DotNetFramework40RegKey = @"Software\Microsoft\NET Framework Setup\NDP\v4\Full";

        /// <summary>
        /// Office user settings Path.
        /// </summary>
        private const string OfficeUserSettings = @"Software\Microsoft\Office\";

        /// <summary>
        /// User settings name.
        /// </summary>
        private const string UserSettings = @"\User Settings";

        /// <summary>
        /// Key for the excel add-in name.
        /// </summary>
        private const string ExcelAddinName = "Microsoft.Research.Wwt.Excel.Addin";

        /// <summary>
        /// The addin path for powerpoint.
        /// </summary>
        private const string ApplicationName = @"Software\Microsoft\Office\Excel\Addins\";

        /// <summary>
        /// Registry key name for Description.
        /// </summary>
        private const string Description = "Description";

        /// <summary>
        /// Description of the Add-in.
        /// </summary>
        private const string AddInDescription = "WorldWide Telescope Add-in for Excel - Excel add-in created with Visual Studio Tools for Office";

        /// <summary>
        /// Registry key name for FriendlyName.
        /// </summary>
        private const string FriendlyName = "FriendlyName";

        /// <summary>
        /// FriendlyName of the Add-in.
        /// </summary>
        private const string AddInFriendlyName = "WorldWide Telescope Add-in for Excel";

        /// <summary>
        /// Registry key name for LoadBehavior.
        /// </summary>
        private const string LoadBehavior = "LoadBehavior";

        /// <summary>
        /// LoadBehavior of the Add-in.
        /// </summary>
        private const int AddInLoadBehavior = 3;

        /// <summary>
        /// Registry key name for Manifest.
        /// </summary>
        private const string Manifest = "Manifest";

        /// <summary>
        /// Manifest extension of the Add-in.
        /// </summary>
        private const string ManifestVSTOExtension = ".vsto|vstolocal";

        /// <summary>
        /// Gets the processor Architecture of the machine.
        /// </summary>
        private static ProcessorType ProcessorArchitecture
        {
            get
            {
                ProcessorType processorType = ProcessorType.x86;
                if (Environment.Is64BitOperatingSystem == true)
                {
                    processorType = ProcessorType.x64;
                }

                return processorType;
            }
        }

        /// <summary>
        /// Gets or sets the Installation Path of the AddIn.
        /// </summary>
        private static string AddInPath { get; set; }

        private static Session Sessions { get; set; }

        /// <summary>
        /// Registers the add-in in the user settings under HKLM\Software\Office\version\User Settings.
        /// </summary>
        /// <param name="session">Session object</param>
        /// <returns>The result after executing the custom action.</returns>
        [CustomAction]
        public static ActionResult RegisterAddin(Session session)
        {
            Sessions = session;
            ActionResult result = ActionResult.Success;
            if (session != null)
            {
                string isOffice2007 = session.CustomActionData["HASEXCEL2007"];
                string isOffice2010 = session.CustomActionData["HASEXCEL2010"];
                string isOffice2010x64 = session.CustomActionData["HASEXCEL2010X64"];
                AddInPath = session.CustomActionData["INSTALLDIR"];

                Sessions.Log("devin 1");
                if (!String.IsNullOrEmpty(isOffice2010))
                {
                    Sessions.Log("devin 2");
                    result = UpdateRegistry(Office2010Version, true, false);
                }

                if (result == ActionResult.Success)
                {
                    if (!String.IsNullOrEmpty(isOffice2007))
                    {
                        Sessions.Log("devin 3");
                        result = UpdateRegistry(Office2007Version, true, false);
                    }
                }

                if (result == ActionResult.Success)
                {
                    if (!String.IsNullOrEmpty(isOffice2010x64) && ProcessorArchitecture == ProcessorType.x64)
                    {
                        Sessions.Log("devin 4");
                        result = UpdateRegistry(Office2010Version, true, true);
                    }
                }
                Sessions.Log("devin 5");
            }

            return result;
        }

        /// <summary>
        ///  Unregisters the add-in in the user settings under HKLM\Software\Office\version\User Settings.
        /// </summary>
        /// <param name="session">Session object.</param>
        /// <returns>The result after executing the custom action.</returns>
        [CustomAction]
        public static ActionResult UnregisterAddin(Session session)
        {
            Sessions = session;
            ActionResult result = ActionResult.Success;
            if (session != null)
            {
                string isOffice2007 = session.CustomActionData["HASEXCEL2007"];
                string isOffice2010 = session.CustomActionData["HASEXCEL2010"];
                string isOffice2010x64 = session.CustomActionData["HASEXCEL2010X64"];

                if (!String.IsNullOrEmpty(isOffice2007))
                {
                    result = UpdateRegistry(Office2007Version, false, false);
                }

                if (result == ActionResult.Success)
                {
                    if (!String.IsNullOrEmpty(isOffice2010))
                    {
                        result = UpdateRegistry(Office2010Version, false, false);
                    }
                }

                if (result == ActionResult.Success)
                {
                    if (!String.IsNullOrEmpty(isOffice2010x64) && ProcessorArchitecture == ProcessorType.x64)
                    {
                        result = UpdateRegistry(Office2010Version, false, true);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Verify that prerequisites have been installed
        /// </summary>
        /// <param name="session">The session variable for the custom action</param>
        /// <returns>ActionResult instance</returns>
        [CustomAction]
        public static ActionResult ValidatePreRequisiteInstallationState(Session session)
        {
            Sessions = session;
            RegistryKey regKey;
            RegistryKey baseKey;
            ActionResult actionResult = ActionResult.Success;
            if (session != null)
            {
                session["HASEXCEL"] = string.Empty;
                session["VSTO40INSTALLED"] = string.Empty;
                session["NETFRAMEWORK40FULL"] = string.Empty;
                session["PREREQUISITEMISSING"] = string.Empty;

                try
                {
                    // 32 bit registry view. 
                    // This will return WOW6432Node HKLM in case of 64 bit machines. In case of
                    // of 32 bit machines, default HKLM will be returned.
                    using (baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32))
                    {
                        using (regKey = baseKey.OpenSubKey(CustomActions.Excel2007RegKey))
                        {
                            if (regKey != null)
                            {
                                string excelPath = (string)regKey.GetValue("Path");
                                if (!string.IsNullOrEmpty(excelPath))
                                {
                                    session["HASEXCEL"] = "1";
                                }
                            }
                        }

                        using (regKey = baseKey.OpenSubKey(CustomActions.Excel2010RegKey))
                        {
                            if (regKey != null)
                            {
                                string excelPath = (string)regKey.GetValue("Path");
                                if (!string.IsNullOrEmpty(excelPath))
                                {
                                    session["HASEXCEL"] = "1";
                                }
                            }
                        }

                        using (regKey = baseKey.OpenSubKey(CustomActions.Excel2013RegKey))
                        {
                            if (regKey != null)
                            {
                                string excelPath = (string)regKey.GetValue("Path");
                                if (!string.IsNullOrEmpty(excelPath))
                                {
                                    session["HASEXCEL"] = "1";
                                }
                            }
                        }
                    }

                    if (string.IsNullOrEmpty(session["HASEXCEL"]))
                    {
                        // 64 bit registry view.
                        // This will return default HKLM in case of both 64 bit machines and 32 bit machines.
                        using (baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                        {
                            using (regKey = baseKey.OpenSubKey(CustomActions.Excel2007RegKey))
                            {
                                if (regKey != null)
                                {
                                    string excelPath = (string)regKey.GetValue("Path");
                                    if (!string.IsNullOrEmpty(excelPath))
                                    {
                                        session["HASEXCEL"] = "1";
                                    }
                                }
                            }

                            using (regKey = baseKey.OpenSubKey(CustomActions.Excel2010RegKey))
                            {
                                if (regKey != null)
                                {
                                    string excelPath = (string)regKey.GetValue("Path");
                                    if (!string.IsNullOrEmpty(excelPath))
                                    {
                                        session["HASEXCEL"] = "1";
                                    }
                                }
                            }

                            using (regKey = baseKey.OpenSubKey(CustomActions.Excel2013RegKey))
                            {
                                if (regKey != null)
                                {
                                    string excelPath = (string)regKey.GetValue("Path");
                                    if (!string.IsNullOrEmpty(excelPath))
                                    {
                                        session["HASEXCEL"] = "1";
                                    }
                                }
                            }
                        }
                    }

                    using (regKey = Registry.LocalMachine.OpenSubKey(CustomActions.VSTO2010RegKey))
                    {
                        if (regKey != null)
                        {
                            string vstoVersion = (string)regKey.GetValue("Version");
                            if (!string.IsNullOrEmpty(vstoVersion))
                            {
                                session["VSTO40INSTALLED"] = "1";
                            }
                        }
                    }

                    using (regKey = Registry.LocalMachine.OpenSubKey(CustomActions.DotNetFramework40RegKey))
                    {
                        if (regKey != null)
                        {
                            object installedState = (object)regKey.GetValue("Install");
                            if (installedState != null && installedState.ToString() == "1")
                            {
                                session["NETFRAMEWORK40FULL"] = "1";
                            }
                        }
                    }
                }
                catch (ArgumentNullException)
                {
                    actionResult = ActionResult.Failure;
                }
                catch (ObjectDisposedException)
                {
                    actionResult = ActionResult.Failure;
                }
                catch (ArgumentException)
                {
                    actionResult = ActionResult.Failure;
                }
                catch (SecurityException)
                {
                    actionResult = ActionResult.Failure;
                }
                catch (UnauthorizedAccessException)
                {
                    actionResult = ActionResult.Failure;
                }
                catch (IOException)
                {
                    actionResult = ActionResult.Failure;
                }

                if (string.IsNullOrEmpty(session["HASEXCEL"])
                    || string.IsNullOrEmpty(session["VSTO40INSTALLED"])
                    || string.IsNullOrEmpty(session["NETFRAMEWORK40FULL"]))
                {
                    session["PREREQUISITEMISSING"] = "1";
                }
            }

            return actionResult;
        }

        /// <summary>
        /// Replace string method
        /// </summary>
        /// <param name="session">The session variable for the custom action</param>
        /// <returns>ActionResult instance</returns>
        [CustomAction]
        public static ActionResult String_Replace(Session session)
        {
            Sessions = session;
            ActionResult result = ActionResult.Success;
            if (session != null)
            {
                string excelPath = session["EXCELPATH"];
                string excel2013 = session["HASEXCEL2013"];
                string excel201364 = session["HASEXCEL2013X64"];
                //if system is Office 2013 - 32 / 64 bit
                //remove the following characters which the installer added to it - ' (x86)'
                if (excel2013.Length > 0 || excel201364.Length > 0)
                {
                    excelPath = excelPath.Replace(@" (x86)\", @"\");
                    //replace modified string in sessions
                    session["EXCELPATH"] = excelPath;
                }
            }
            return result;
        }

        /// <summary>
        /// Helps to register and unregister the addin.
        /// </summary>
        /// <param name="officeVersion">Installed Office version.</param>
        /// <param name="register">Indicates if it Register/Un-register operation.</param>
        /// <param name="isOffice64Bit">Indicates if office is 64 bit.</param>
        /// <returns>The result after reading/writing into the registry.</returns>
        private static ActionResult UpdateRegistry(string officeVersion, bool register, bool isOffice64Bit)
        {
            RegistryKey userSettingsKey = null;
            RegistryKey instructionKey = null;
            RegistryKey deleteKey = null;
            RegistryKey baseKey = null;
            ActionResult result = ActionResult.Success;
            string subKey = register ? "DELETE" : "CREATE";

            try
            {
                var userSettingsLocation = new StringBuilder();
                //userSettingsLocation.Append(Path.Combine(OfficeUserSettings, officeVersion));
                //userSettingsLocation.Append(UserSettings);

                userSettingsLocation.Append(ApplicationName);

                Sessions.Log("devin 6 : " + ApplicationName);

                userSettingsKey = GetKey(isOffice64Bit, userSettingsLocation, baseKey);

                if (userSettingsKey == null)
                {
                    Sessions.Log("devin 6a : " + ApplicationName);
                    var officeKey = GetKey(isOffice64Bit, new StringBuilder(OfficeUserSettings), baseKey);

                    if (officeKey != null)
                    {
                        Sessions.Log("devin 6s : " + ApplicationName);
                        RegistryKey excelKey = officeKey.OpenSubKey(@"Excel\Addins\", true);
                        if (excelKey == null)
                        {
                            Sessions.Log("devin 6d : " + ApplicationName);
                            excelKey = officeKey.CreateSubKey(@"Excel\Addins\");
                        }

                        Sessions.Log("devin 6f : " + ApplicationName);
                        userSettingsKey = GetKey(isOffice64Bit, userSettingsLocation, baseKey);
                    }
                    else
                    {
                        return ActionResult.Failure;
                    }
                }

                if (userSettingsKey != null)
                {
                    Sessions.Log("devin 7 : " + userSettingsKey.ToString());

                    instructionKey = userSettingsKey.OpenSubKey(ExcelAddinName, true);
                    if (instructionKey == null)
                    {
                        Sessions.Log("devin 8 : ");
                        instructionKey = userSettingsKey.CreateSubKey(ExcelAddinName);
                    }
                    else
                    {
                        Sessions.Log("devin 9 : " + instructionKey.ToString());

                        // Remove the Delete instruction
                        if (instructionKey.GetSubKeyNames().Where(name => name.Equals(subKey)).Count() > 0)
                        {
                            instructionKey.DeleteSubKeyTree(subKey);
                        }

                        if (!register)
                        {
                            string instructionString = @"DELETE\" + ApplicationName + @"\" + ExcelAddinName;
                            deleteKey = instructionKey.CreateSubKey(instructionString);
                        }

                        Sessions.Log("devin 10 : ");
                    }

                    if (isOffice64Bit)
                    {
                        Sessions.Log("devin 11 : ");
                        UpdateRegistryx64(instructionKey, register);
                    }

                    Sessions.Log("devin 12 : ");
                    IncrementCounter(instructionKey);
                    Sessions.Log("devin 13 : ");
                }
                else
                {
                    result = ActionResult.Failure;
                }
            }
            catch (ArgumentNullException)
            {
                result = ActionResult.Failure;
            }
            catch (ArgumentException)
            {
                result = ActionResult.Failure;
            }
            catch (ObjectDisposedException)
            {
                result = ActionResult.Failure;
            }
            catch (SecurityException)
            {
                result = ActionResult.Failure;
            }
            catch (UnauthorizedAccessException)
            {
                result = ActionResult.Failure;
            }
            catch (IOException)
            {
                result = ActionResult.Failure;
            }
            finally
            {
                if (deleteKey != null)
                {
                    deleteKey.Close();
                }

                if (instructionKey != null)
                {
                    instructionKey.Close();
                }

                if (userSettingsKey != null)
                {
                    userSettingsKey.Close();
                }

                if (baseKey != null)
                {
                    baseKey.Close();
                }
            }

            return result;
        }

        private static RegistryKey GetKey(bool isOffice64Bit, StringBuilder userSettingsLocation, RegistryKey baseKey)
        {
            RegistryKey userSettingsKey = null;
            // Write to the 64-bit registry from a 32-bit installer.
            if (isOffice64Bit)
            {
                //Sessions.Log("devin 6.1 : " + ApplicationName);
                baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
                if (baseKey != null)
                {
                    userSettingsKey = baseKey.OpenSubKey(userSettingsLocation.ToString(), true);
                }
            }
            else
            {
                Sessions.Log("devin 6.2 : " + ApplicationName);
                userSettingsKey = Registry.LocalMachine.OpenSubKey(userSettingsLocation.ToString(), true);
            }
            return userSettingsKey;
        }

        /// <summary>
        /// Helps to register and unregister the addin for 64 bit office.
        /// </summary>
        /// <param name="regKey">Registry key for the Office version.</param>
        /// <param name="register">Indicates if it Register/Un-register operation.</param>
        private static void UpdateRegistryx64(RegistryKey regKey, bool register)
        {
            RegistryKey instructionKey = null;

            if (regKey != null)
            {
                if (register)
                {
                    string addinManifest = Path.Combine(AddInPath ?? string.Empty, ExcelAddinName) + ManifestVSTOExtension;
                    string subKey = "CREATE";
                    instructionKey = regKey.OpenSubKey(subKey, true);

                    if (instructionKey == null)
                    {
                        instructionKey = regKey.CreateSubKey(subKey + @"\" + ApplicationName + @"\" + ExcelAddinName);

                        if (register)
                        {
                            instructionKey.SetValue(Description, AddInDescription);
                            instructionKey.SetValue(FriendlyName, AddInFriendlyName);
                            instructionKey.SetValue(LoadBehavior, AddInLoadBehavior);
                            instructionKey.SetValue(Manifest, addinManifest);
                        }
                    }

                    regKey.SetValue(Description, AddInDescription);
                    regKey.SetValue(FriendlyName, AddInFriendlyName);
                    regKey.SetValue(LoadBehavior, AddInLoadBehavior);
                    regKey.SetValue(Manifest, addinManifest);
                }
            }
        }

        /// <summary>        
        /// Responsible for correctly updating the Count registry value under HKLM.
        /// </summary>
        /// <param name="instructionKey">The registry key.</param>
        private static void IncrementCounter(RegistryKey instructionKey)
        {
            var count = 1;
            var value = instructionKey.GetValue("Count");

            if (value != null)
            {
                if ((int)value != Int32.MaxValue)
                {
                    count = (int)value + 1;
                }
            }

            instructionKey.SetValue("Count", count);
        }
    }
}
