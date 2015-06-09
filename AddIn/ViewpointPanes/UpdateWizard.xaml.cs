﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// Interaction logic for UpdateWizard.xaml
    /// </summary>
    public partial class UpdateWizard : Window
    {
        public UpdateWizard()
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
