//-----------------------------------------------------------------------
// <copyright file="LayerDropdownProperties.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using Microsoft.Research.Wwt.Excel.Common;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Layer drop down properties
    /// </summary>
    public class LayerDropdownProperties : PropertyChangeBase
    {
        #region Private Properties
        private bool isSelectable;
        private bool isDeleted;
        #endregion

        /// <summary>
        /// Gets or sets a value indicating whether the reference frame/group/layer is selectable
        /// </summary>
        public bool IsSelectable
        {
            get 
            {
                return this.isSelectable; 
            }
            set
            {
                this.isSelectable = value;
                OnPropertyChanged("IsSelectable"); 
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the reference frame/group gets deleted.
        /// </summary>
        public bool IsDeleted
        {
            get
            {
                return this.isDeleted;
            }
            set
            {
                this.isDeleted = value;
                OnPropertyChanged("IsDeleted"); 
            }
        }

        /// <summary>
        /// Gets or sets Layer ID
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets Layer name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets Layer map type
        /// </summary>
        public LayerMapType LayerType { get; set; }
    }
}
