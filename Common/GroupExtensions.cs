//-----------------------------------------------------------------------
// <copyright file="GroupExtensions.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Text;
using System.Xml;
namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// Extension class for Group
    /// </summary>
    public static class GroupExtensions
    {
        /// <summary>
        /// This extension method checks if the layer is of type planet.
        /// </summary>
        /// <param name="group">
        /// Group details
        /// </param>
        /// <returns>
        /// True if the layer is of type Planets;Otherwise false.
        /// </returns>
        public static bool IsPlanet(this Group group)
        {
            return group != null && group.Path.StartsWith(Constants.SunReferencePath, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// This extension method gets reference frame name from the Group.
        /// </summary>
        /// <param name="group">
        /// Group details
        /// </param>
        /// <returns>
        ///  Reference frame name from the Group
        /// </returns>
        public static string GetReferenceFrame(this Group group)
        {
            string referenceFrame = string.Empty;
            if (group != null)
            {
                if (group.GroupType == GroupType.ReferenceFrame)
                {
                    referenceFrame = group.Name;
                }
                else
                {
                    referenceFrame = group.Parent.GetReferenceFrame();
                }
            }

            return referenceFrame;
        }

        /// <summary>
        /// Extension method to serialize the Group
        /// </summary>
        /// <param name="group">group object</param>
        /// <returns>serialized string</returns>
        public static string Serialize(this Group group)
        {
            StringBuilder serializedString = new StringBuilder();
            if (group != null)
            {
                try
                {
                    using (var writer = XmlWriter.Create(serializedString))
                    {
                        var serializer = new DataContractSerializer(typeof(Group), Common.Constants.GroupXmlRootName, Common.Constants.GroupXmlNamespace);
                        if (writer != null)
                        {
                            serializer.WriteObject(writer, group);
                        }
                    }
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
        /// Extension method to de-serialize the string into a Group
        /// </summary>
        /// <param name="group">Group object</param>
        /// <param name="xmlContent">xml content</param>
        /// <returns>populated Group object</returns>
        public static Group Deserialize(this Group group, string xmlContent)
        {
            try
            {
                using (var stringReader = new StringReader(xmlContent))
                {
                    var reader = XmlReader.Create(stringReader);
                    {
                        var serializer = new DataContractSerializer(typeof(Group), Common.Constants.GroupXmlRootName, Common.Constants.GroupXmlNamespace);
                        group = (Group)serializer.ReadObject(reader, true);
                    }
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

            return group;
        }

        /// <summary>
        /// This function is used to retrieve the default group for earth.
        /// </summary>
        /// <param name="groups">
        /// Collection of WWT groups.
        /// </param>
        /// <returns>
        /// Default Group instance.
        /// </returns>
        public static Group GetDefaultEarthGroup(this ICollection<Group> groups)
        {
            Group defaultGroup = SearchGroup(groups, Common.Constants.EarthReferenceFrame);
            if (defaultGroup == null)
            {
                Group sun = new Group(Common.Constants.SunFrameName, GroupType.ReferenceFrame, null)
                {
                    IsDeleted = true
                };

                defaultGroup = new Group(Common.Constants.EarthReferenceFrame, GroupType.ReferenceFrame, sun)
                {
                    IsDeleted = true
                };
            }

            return defaultGroup;
        }

        /// <summary>
        /// This function is used to retrieve the default group for sky.
        /// </summary>
        /// <param name="groups">
        /// Collection of WWT groups.
        /// </param>
        /// <returns>
        /// Default Group instance.
        /// </returns>
        public static Group GetDefaultSkyGroup(this ICollection<Group> groups)
        {
            Group defaultGroup = SearchGroup(groups, Common.Constants.SkyReferenceFrame);
            if (defaultGroup == null)
            {
                defaultGroup = new Group(Common.Constants.SkyReferenceFrame, GroupType.ReferenceFrame, null)
                 {
                     IsDeleted = true
                 };
            }

            return defaultGroup;
        }

        /// <summary>
        /// This function is used to search for the given group in the WWT Groups list.
        /// </summary>
        /// <param name="wwtGroups">
        /// List of WWT groups.
        /// </param>
        /// <param name="groupName">
        /// Name of the Group.
        /// </param>
        /// <returns>
        /// A Group that represents the group name and path.
        /// </returns>
        public static Group SearchGroup(this ICollection<Group> wwtGroups, string groupName)
        {
            Group result = null;
            if (wwtGroups != null)
            {
                foreach (Group group in wwtGroups)
                {
                    result = (string.CompareOrdinal(groupName, group.Name) == 0) ? group : SearchGroup(group.Children, groupName);

                    if (result != null)
                    {
                        break;
                    }
                }
            }

            return result;
        }
    }
}
