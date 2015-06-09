//-----------------------------------------------------------------------
// <copyright file="RelayCommand.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Windows.Input;

namespace Microsoft.Research.Wwt.Excel.Addin
{
    /// <summary>
    /// This class is the base class that implements the ICommand interface.
    /// Two methods of the ICommand are implemented as Abstract so that they are mandated to be overridden in derived class.
    /// </summary>
    public abstract class RelayCommand : ICommand
    {
        #region ICommand Members

        /// <summary>
        /// This is the handler for there binding system to refresh the CanExecute value.
        /// </summary>
        public event EventHandler CanExecuteChanged;

        /// <summary>
        /// This method should return true if the command can be executed
        /// </summary>
        /// <param name="parameter">Command parameter</param>
        /// <returns>True / false if the command can be executed</returns>
        public virtual bool CanExecute(object parameter)
        {
            return true;
        }

        /// <summary>
        /// This method is called when the control to which a command is bound
        /// is actuated by the user
        /// </summary>
        /// <param name="parameter">Command parameter</param>
        public abstract void Execute(object parameter);

        /// <summary>
        /// Raise the CanExecuteChanged event.
        /// </summary>
        public void OnRaiseCanExecuteChanged()
        {
            if (this.CanExecuteChanged != null)
            {
                this.CanExecuteChanged(this, new EventArgs());
            }
        }

        #endregion // ICommand Members
    }
}
