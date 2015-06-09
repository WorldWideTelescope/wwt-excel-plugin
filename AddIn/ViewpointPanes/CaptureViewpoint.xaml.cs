//-----------------------------------------------------------------------
// <copyright file="CaptureViewpoint.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Windows;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Interaction logic for CaptureViewpoint.xaml
    /// </summary>
    public partial class CaptureViewpoint : Window
    {
        public CaptureViewpoint()
        {
            this.InitializeComponent();
            this.Loaded += new RoutedEventHandler(OnCaptureViewpointLoaded);
        }

        internal void OnRequestClose(object sender, EventArgs e)
        {
            this.DialogResult = true; 
        }

        private void OnCaptureViewpointLoaded(object sender, RoutedEventArgs e)
        {
            this.viewpointName.Focus();
            this.viewpointName.SelectAll();
        }
    }
}