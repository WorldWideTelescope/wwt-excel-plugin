//-----------------------------------------------------------------------
// <copyright file="LayerManagerPaneHost.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using System.Windows.Forms;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Interaction logic for LayerManagerPaneHost.
    /// </summary>
    public partial class LayerManagerPaneHost : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the LayerManagerPaneHost class.
        /// </summary>
        public LayerManagerPaneHost()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Gets LayerManagerPane instance 
        /// </summary>
        public LayerManagerPane LayerManagerPane
        {
            get { return this.layerManagerPane; }
        }
    }
}
