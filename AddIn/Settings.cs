//-----------------------------------------------------------------------
// <copyright file="Settings.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Configuration;
namespace Microsoft.Research.Wwt.Excel.Addin.Properties
{
    // This class allows you to handle specific events on the settings class:
    //  The SettingChanging event is raised before a setting's value is changed.
    //  The PropertyChanged event is raised after a setting's value is changed.
    //  The SettingsLoaded event is raised after the setting values are loaded.
    //  The SettingsSaving event is raised before the setting values are saved.
    [SettingsProvider(typeof(CustomSettingsProvider))]
    internal sealed partial class Settings
    {
        public Settings()
        {
        }
    }
}
