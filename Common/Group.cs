//-----------------------------------------------------------------------
// <copyright file="Group.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.ObjectModel;
using System.Runtime.Serialization;

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// This contains the information about the group.
    /// </summary>
    [DataContract]
    public class Group : IEquatable<Group>
    {
        /// <summary>
        /// Initializes a new instance of the Group class.
        /// </summary>
        /// <param name="name">
        /// Name of group.
        /// </param>
        /// <param name="type">
        /// Type of the group.
        /// </param>
        /// <param name="parent">
        /// Parent of the group.
        /// </param>
        public Group(string name, GroupType type, Group parent)
        {
            this.Name = name;
            this.GroupType = type;
            this.Parent = parent;
            if (parent != null)
            {
                this.Path = parent.Path;
            }
            this.Path += "/" + name;

            this.Children = new Collection<Group>();
            this.LayerIDs = new Collection<string>();
        }

        /// <summary>
        /// Gets the name of the group.
        /// </summary>
        [DataMember]
        public string Name
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the type of the group.
        /// </summary>
        [DataMember]
        public GroupType GroupType
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the parent of the group.
        /// </summary>
        [DataMember]
        public Group Parent
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the path of the group.
        /// </summary>
        [DataMember]
        public string Path
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the children of the group.
        /// </summary>
        public Collection<Group> Children
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets all the layer id's of the group.
        /// </summary>
        public Collection<string> LayerIDs
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets or sets a value indicating whether the group is deleted in WWT.
        /// </summary>
        public bool IsDeleted
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets a value indicating whether the group is expanded or not
        /// </summary>
        public bool IsExpanded
        {
            get;
            set;
        }

        /// <summary>
        /// Returns a System.String that represents the current group.
        /// </summary>
        /// <returns>
        /// A System.String that represents the current group.
        /// </returns>
        public override string ToString()
        {
            return string.Format(System.Globalization.CultureInfo.CurrentCulture, "Name = {0} , Path = {1}", this.Name, this.Path);
        }

        #region IEquatable

        /// <summary>
        /// Determines whether the Groups are equal.
        /// </summary>
        /// <param name="other">
        /// The group to be compared with
        /// </param>
        /// <returns>
        /// true if the Groups are equal; otherwise, false.
        /// </returns>
        public bool Equals(Group other)
        {
            if (other == null)
            {
                return false;
            }
            return (Name.Equals(other.Name, StringComparison.OrdinalIgnoreCase) && Path.Equals(other.Path, StringComparison.OrdinalIgnoreCase));
        }

        #endregion
    }
}
