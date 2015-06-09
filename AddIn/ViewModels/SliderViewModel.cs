//-----------------------------------------------------------------------
// <copyright file="SliderViewModel.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// View Model for slider
    /// </summary>
    public class SliderViewModel : PropertyChangeBase
    {
        #region Private properties

        /// <summary>
        /// maximum Number
        /// </summary>
        private double maximumNum;

        /// <summary>
        /// minimum Number
        /// </summary>
        private double minimumNum;

        /// <summary>
        /// tick Frequency
        /// </summary>
        private double tickFrequency;

        /// <summary>
        /// Collection of ticks
        /// </summary>
        private Collection<double> sliderTicks;

        /// <summary>
        /// tool tip value to be shown
        /// </summary>
        private string tooltipValue;

        /// <summary>
        /// Selected slider value
        /// </summary>
        private double selectedSliderValue;

        #endregion Private properties

        /// <summary>
        /// Initializes a new instance of the SliderViewModel class
        /// </summary>
        /// <param name="sliderTicks">collection of ticks</param>
        public SliderViewModel(Collection<double> sliderTicks)
        {
            if (sliderTicks != null)
            {
                this.sliderTicks = sliderTicks;
                this.SetSliderValues();
            }
        }

        #region Public properties

        /// <summary>
        /// Gets maximum value
        /// </summary>
        public double MaximumValue
        {
            get
            {
                return this.maximumNum;
            }
        }

        /// <summary>
        /// Gets slide tick values in a string format
        /// </summary>
        public string SliderTicksValues
        {
            get
            {
                return String.Join(",", this.sliderTicks.ToArray());
            }
        }

        /// <summary>
        /// Gets slide ticks as collection
        /// </summary>
        public Collection<double> SliderTicks
        {
            get
            {
                return this.sliderTicks;
            }
        }

        /// <summary>
        /// Gets or sets tool tip value for slider tick
        /// </summary>
        public string ToolTipValue
        {
            get
            {
                return this.tooltipValue;
            }
            set
            {
                this.tooltipValue = value;
            }
        }

        /// <summary>
        /// Gets minimum value
        /// </summary>
        public double MinimumValue
        {
            get
            {
                return this.minimumNum;
            }
        }

        /// <summary>
        /// Gets tick frequency value
        /// </summary>
        public double TickFrequency
        {
            get
            {
                return this.tickFrequency;
            }
        }

        /// <summary>
        /// Gets or sets selected slider value
        /// </summary>
        public double SelectedSliderValue
        {
            get
            {
                return this.selectedSliderValue;
            }
            set
            {
                this.selectedSliderValue = value;
                OnPropertyChanged("SelectedSliderValue");
            }
        }

        #endregion Public properties

        /// <summary>
        /// Sets values for slider 
        /// </summary>
        private void SetSliderValues()
        {
            this.minimumNum = this.sliderTicks.Min();
            this.maximumNum = this.sliderTicks.Max();
            this.tickFrequency = this.sliderTicks.ElementAtOrDefault<double>(1) - this.sliderTicks.First<double>();
        }
    }
}
