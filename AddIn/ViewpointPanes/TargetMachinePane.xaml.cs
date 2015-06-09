//-----------------------------------------------------------------------
// <copyright file="TargetMachinePane.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Windows;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Interaction logic for TargetMachinePane.xaml
    /// </summary>
    public partial class TargetMachinePane : Window
    {
        public TargetMachinePane()
        {
            this.InitializeComponent();
            this.Loaded += new RoutedEventHandler(OnTargetMachinePaneLoaded);
        }

        internal void OnRequestClose(object sender, EventArgs e)
        {
            this.DialogResult = true;
        }

        private void OnTargetMachinePaneLoaded(object sender, RoutedEventArgs e)
        {
            this.targetMachine.Focus();
            this.targetMachine.SelectAll();
        }
    }
}