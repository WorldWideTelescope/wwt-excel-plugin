//-----------------------------------------------------------------------
// <copyright file="ViewpointMapExtensions.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.IO;
using System.Runtime.Serialization;
using System.Text;
using System.Xml;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Extension methods on ViewpointMap
    /// </summary>
    internal static class ViewpointMapExtensions
    {
        /// <summary>
        /// Extension method to serialize the Viewpoint Map
        /// </summary>
        /// <param name="viewpointMap">Viewpoint Map object</param>
        /// <returns>serialized string</returns>
        internal static string Serialize(this ViewpointMap viewpointMap)
        {
            StringBuilder serializedString = new StringBuilder();
            if (viewpointMap != null)
            {
                try
                {
                    using (var writer = XmlWriter.Create(serializedString))
                    {
                        var serializer = new DataContractSerializer(typeof(ViewpointMap), Common.Constants.ViewpointMapRootName, Common.Constants.ViewpointMapXmlNamespace);
                        if (writer != null)
                        {
                            serializer.WriteObject(writer, viewpointMap);
                        }
                    }
                }
                catch (ArgumentNullException ex)
                {
                    Logger.LogException(ex);
                }
                catch (InvalidDataContractException ex)
                {
                    Logger.LogException(ex);
                }
                catch (SerializationException ex)
                {
                    Logger.LogException(ex);
                }
            }

            return serializedString.ToString();
        }

        /// <summary>
        /// Extension method to de-serialize the string into a viewpoint Map
        /// </summary>
        /// <param name="viewpointMap">viewpointMap object</param>
        /// <param name="xmlContent">xml content</param>
        /// <returns>populated viewpointMap object</returns>
        internal static ViewpointMap Deserialize(this ViewpointMap viewpointMap, string xmlContent)
        {
            if (viewpointMap != null)
            {
                using (var stringReader = new StringReader(xmlContent))
                {
                    try
                    {
                        var reader = XmlReader.Create(stringReader);
                        {
                            var serializer = new DataContractSerializer(typeof(ViewpointMap), Common.Constants.ViewpointMapRootName, Common.Constants.ViewpointMapXmlNamespace);
                            viewpointMap = (ViewpointMap)serializer.ReadObject(reader, true);
                        }
                    }
                    catch (ArgumentNullException ex)
                    {
                        Logger.LogException(ex);
                    }
                    catch (SerializationException ex)
                    {
                        Logger.LogException(ex);
                    }
                }
            }

            return viewpointMap;
        }
    }
}
