//-----------------------------------------------------------------------
// <copyright file="CustomSettingsProvider.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// This class is custom setting provider. 
    /// </summary>
    public class CustomSettingsProvider : SettingsProvider
    {
        private XElement settingXML;

        /// <summary>
        /// Gets or sets the name of the currently running application.
        /// </summary>
        public override string ApplicationName
        {
            get
            {
                return (System.Reflection.Assembly.GetExecutingAssembly().GetName().Name);
            }
            set
            {
                // Do nothing.
            }
        }

        /// <summary>
        /// Gets the default app setting path.
        /// </summary>
        /// <returns>
        /// </returns>
        private static string GetAppSettingsPath
        {
            get
            {
                string appsettingPath = Path.Combine(Application.LocalUserAppDataPath, Constants.ConfigFileName);

                if (!Directory.Exists(Application.LocalUserAppDataPath))
                {
                    Directory.CreateDirectory(Application.LocalUserAppDataPath);
                }

                return appsettingPath;
            }
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "We should not throw any exception from custom setting providers.")]
        private XElement SettingXML
        {
            get
            {
                try
                {
                    settingXML = XElement.Load(GetAppSettingsPath);
                }
                catch
                {
                    settingXML = new XElement(Constants.SettingNodeName);
                }

                return settingXML;
            }
        }

        /// <summary>
        /// Initializes the provider.
        /// </summary>
        /// <param name="name">
        /// The friendly name of the provider.
        /// </param>
        /// <param name="config">
        /// A collection of the name/value pairs representing the provider-specific attributes
        /// specified in the configuration for this provider.
        /// </param>
        public override void Initialize(string name, System.Collections.Specialized.NameValueCollection config)
        {
            base.Initialize(this.ApplicationName, config);
        }

        /// <summary>
        /// Returns the collection of settings property values for the specified application
        ///     instance and settings property group.
        /// </summary>
        /// <param name="context">
        /// A System.Configuration.SettingsContext describing the current application use.
        /// </param>
        /// <param name="collection">
        ///  A System.Configuration.SettingsPropertyCollection containing the settings
        ///  property group whose values are to be retrieved.
        ///  </param>
        /// <returns>
        /// A System.Configuration.SettingsPropertyValueCollection containing the values
        /// for the specified settings property group.
        /// </returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "We should not throw any exceptions from custom setting providers.")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2122:DoNotIndirectlyExposeMethodsWithLinkDemands", Justification = "This class needs link demand.")]
        public override SettingsPropertyValueCollection GetPropertyValues(SettingsContext context, SettingsPropertyCollection collection)
        {
            // Create new collection of values
            SettingsPropertyValueCollection values = new SettingsPropertyValueCollection();

            if (collection != null)
            {
                try
                {
                    // Get the latest config setting from Local app user data folder.
                    XElement configSettings = this.SettingXML ?? new XElement(Constants.SettingNodeName);

                    // Iterate through the settings to be retrieved
                    foreach (SettingsProperty setting in collection)
                    {
                        SettingsPropertyValue value = new SettingsPropertyValue(setting);
                        value.IsDirty = false;
                        value.SerializedValue = GetPropertyValue(setting, configSettings);
                        values.Add(value);
                    }
                }
                catch (Exception exception)
                {
                    // Ignore all exceptions from the Custom settings provider.
                    // Refer link :- http://msdn.microsoft.com/en-us/library/8eyb2ct1.aspx
                    Logger.LogException(exception);
                }
            }
            return values;
        }

        /// <summary>
        /// Sets the values of the specified group of property settings.
        /// </summary>
        /// <param name="context">
        /// A System.Configuration.SettingsContext describing the current application usage.
        /// </param>
        /// <param name="collection">
        /// A System.Configuration.SettingsPropertyValueCollection representing the group
        /// of property settings to set.
        /// </param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "We should not throw any exceptions from custom setting providers.")]
        public override void SetPropertyValues(SettingsContext context, SettingsPropertyValueCollection collection)
        {
            if (collection != null)
            {
                try
                {
                    // Get the latest config setting from Local app user data folder.
                    XElement configSettings = this.SettingXML ?? new XElement(Constants.SettingNodeName);

                    // Iterate through the settings to be stored
                    foreach (SettingsPropertyValue propval in collection)
                    {
                        // NOTE: this provider allows setting to both user- and application-scoped
                        // settings. The default provider for ApplicationSettingsBase - 
                        // LocalFileSettingsProvider - is read-only for application-scoped setting. This 
                        // is an example of a policy that a provider may need to enforce for implementation,
                        // security or other reasons.
                        SetPropertyValue(propval, configSettings);
                    }

                    configSettings.Save(GetAppSettingsPath);
                }
                catch (Exception exception)
                {
                    // Ignore all exceptions from the Custom settings provider.
                    // Refer link :- http://msdn.microsoft.com/en-us/library/8eyb2ct1.aspx
                    Logger.LogException(exception);
                }
            }
        }

        /// <summary>
        /// Sets the property value in the app settings xml.
        /// </summary>
        /// <param name="prop">
        /// Property instance.
        /// </param>
        /// <param name="configSettings">
        /// Configuration setting XElement.
        /// </param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2122:DoNotIndirectlyExposeMethodsWithLinkDemands", Justification = "This class needs link demand.")]
        private static void SetPropertyValue(SettingsPropertyValue prop, XElement configSettings)
        {
            var properties = from el in configSettings.Elements(Constants.AddNodeName)
                             where el.Attribute(Constants.KeyAttributeName).Value == prop.Name
                             select el;

            if (properties != null && properties.Count() >= 1)
            {
                properties.First().Attribute(Constants.ValueAttributeName).SetValue(prop.SerializedValue);
            }
            else
            {
                configSettings.Add(new XElement(
                    Constants.AddNodeName,
                    new XAttribute(Constants.KeyAttributeName, prop.Name),
                    new XAttribute(Constants.ValueAttributeName, prop.SerializedValue)));
            }
        }

        /// <summary>
        /// Gets the value of the property.
        /// </summary>
        /// <param name="prop">
        /// Property instance.
        /// </param>
        /// <param name="configSettings">
        /// Configuration setting XElement.
        /// </param>
        /// <returns>
        /// Value of the property.
        /// </returns>
        private static object GetPropertyValue(SettingsProperty prop, XElement configSettings)
        {
            var properties = from el in configSettings.Elements(Constants.AddNodeName)
                             where el.Attribute(Constants.KeyAttributeName).Value == prop.Name
                             select el;

            if (properties != null && properties.Count() >= 1)
            {
                return properties.First().Attribute(Constants.ValueAttributeName).Value;
            }
            else if (prop.DefaultValue != null)
            {
                return prop.DefaultValue;
            }

            return string.Empty;
        }
    }
}
