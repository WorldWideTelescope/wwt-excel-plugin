﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.0
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Microsoft.Research.Wwt.Excel.Common.Properties {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("Microsoft.Research.Wwt.Excel.Common.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to An error has occurred while performing the current operation. Please try again..
        /// </summary>
        internal static string CreateNamedRangeFailure {
            get {
                return ResourceManager.GetString("CreateNamedRangeFailure", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to An error has occurred while performing the current operation. Please try again..
        /// </summary>
        internal static string DefaultErrorMessage {
            get {
                return ResourceManager.GetString("DefaultErrorMessage", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to An error has occurred while performing the current operation. Please try again..
        /// </summary>
        internal static string DownloadUpdatesResponseError {
            get {
                return ResourceManager.GetString("DownloadUpdatesResponseError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to An error has occurred while performing the current operation. Please try again..
        /// </summary>
        internal static string ErrorFromLCAPICall {
            get {
                return ResourceManager.GetString("ErrorFromLCAPICall", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Your system ({0}) is not authorized by the remote WWT client that you are trying to connect to.
        ///
        ///To resolve this issue, from the WWT client
        ///1. Open Remote Access Control Dialog (Settings-&gt;Remote Access Control)
        ///2. Add your IP address ({0}) to the Accept List.
        ///
        ///Once added to the accept list, try again..
        /// </summary>
        internal static string ErrorLCAPIConnectionFailure {
            get {
                return ResourceManager.GetString("ErrorLCAPIConnectionFailure", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to An unexpected error occurred. Restart Microsoft Office Excel and then try again. If the problem continues, reinstall {0}..
        /// </summary>
        internal static string ErrorReadingRegistry {
            get {
                return ResourceManager.GetString("ErrorReadingRegistry", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to WorldWide Telescope Add-in for Excel.
        /// </summary>
        internal static string ProductNameShort {
            get {
                return ResourceManager.GetString("ProductNameShort", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Success.
        /// </summary>
        internal static string ResponseSuccessfulText {
            get {
                return ResourceManager.GetString("ResponseSuccessfulText", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Could not connect to the specified machine.
        /// </summary>
        internal static string RetrieveIPAddressFailure {
            get {
                return ResourceManager.GetString("RetrieveIPAddressFailure", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Timeout.
        /// </summary>
        internal static string TimeoutErrorText {
            get {
                return ResourceManager.GetString("TimeoutErrorText", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to WWT needs to be installed for this operation. Install the latest version of WWT from http://layerscape.org..
        /// </summary>
        internal static string WWTNotInstalledError {
            get {
                return ResourceManager.GetString("WWTNotInstalledError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to WorldWide Telescope (WWT) needs to be open to perform this operation. Please open WWT and try again..
        /// </summary>
        internal static string WWTNotOpenFailure {
            get {
                return ResourceManager.GetString("WWTNotOpenFailure", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to This add-in is not compatible with the version of WWT you’re running. Upgrade to the latest version of WWT from http://layerscape.org..
        /// </summary>
        internal static string WWTOlderVersionError {
            get {
                return ResourceManager.GetString("WWTOlderVersionError", resourceCulture);
            }
        }
    }
}
