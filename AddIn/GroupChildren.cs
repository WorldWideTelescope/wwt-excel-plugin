//-----------------------------------------------------------------------
// <copyright file="GroupChildren.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using System.Collections.ObjectModel;
using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// View model class for the reference frame
    /// </summary>
    public class GroupChildren
    {
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the GroupChildren class
        /// </summary>
        public GroupChildren()
        {
            this.Children = new Collection<GroupChildren>();
            this.Layers = new Collection<Layer>();
            this.AllChildren = new Collection<object>();
        }
        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the group instance.
        /// </summary>
        public Group Group
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the name of the group.
        /// </summary>
        public string Name
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the children of the group
        /// </summary>
        public Collection<GroupChildren> Children
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the children of the group including groups and layers.
        /// </summary>
        public Collection<object> AllChildren
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the layers of the group.
        /// </summary>
        public Collection<Layer> Layers
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets or sets a value indicating whether the group is deleted in WWT or not.
        /// </summary>
        public bool IsDeleted
        {
            get;
            set;
        }
        #endregion
    }
}
