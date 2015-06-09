//-----------------------------------------------------------------------
// <copyright file="FetchClimateView.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation 2011. All rights reserved.
// Interaction logic for FetchClimateView.xaml.
// </copyright>
//-----------------------------------------------------------------------

using System.Windows;
using System;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Interaction logic for FetchClimateView.xaml.
    /// </summary>
    public partial class FetchClimateView : Window
    {
        /// <summary>
        /// Default constuctor.
        /// </summary>
        public FetchClimateView()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Dialog close event handler.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        internal void OnRequestClose(object sender, EventArgs e)
        {           
            this.Close();
        }
    }
}
