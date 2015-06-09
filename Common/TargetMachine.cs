//-----------------------------------------------------------------------
// <copyright file="TargetMachine.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation 2011. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Net;
using System.Net.Sockets;
using System.Collections.ObjectModel;

namespace Microsoft.Research.Wwt.Excel.Common
{
    /// <summary>
    /// This class is responsible for the storing the details of the target machine.
    /// </summary>
    public class TargetMachine
    {
        /// <summary>
        /// Initializes a new instance of the TargetMachine class.
        /// </summary>
        /// <param name="inputValue">
        /// Machine name.
        /// </param>
        public TargetMachine(string inputValue)
        {
            // Validate & throw custom exception
            IPAddress machineAddress;
            IPAddress defaultIpAdress = GetDefaultIp();
            if (string.IsNullOrWhiteSpace(inputValue))
            {
                SetMachineIPIfValid(defaultIpAdress, Constants.Localhost);
            }
            else if (IPAddress.TryParse(inputValue.Trim(), out machineAddress))
            {
                SetMachineIPIfValid(machineAddress, machineAddress.ToString());
            }
            else
            {
                SetMachineIPIfValid(GetIpFromName(inputValue), inputValue);
            }

            if (MachineIP.Equals(defaultIpAdress))
            {
                IsLocalMachine = true;
            }
            else
            {
                IsLocalMachine = false;
            }
        }

        /// <summary>
        /// Initializes a new instance of the TargetMachine class.
        /// </summary>
        public TargetMachine()
        {
            IsLocalMachine = true;
            MachineIP = GetDefaultIp();
            DisplayValue = MachineIP.ToString();
        }

        /// <summary>
        /// Gets the IP address of the local machine.
        /// </summary>
        public static IPAddress DefaultIP
        {
            get { return GetDefaultIp(); }
        }

        /// <summary>
        /// Gets the display value of the target machine.
        /// </summary>
        public string DisplayValue { get; private set; }

        /// <summary>
        /// Gets or sets the IP address of the machine
        /// </summary>
        public IPAddress MachineIP { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the current TargetMachine is local machine or not.
        /// </summary>
        public bool IsLocalMachine { get; set; }

        /// <summary>
        /// This function is used to get the default IP Address.
        /// </summary>
        /// <returns>
        /// Default IP address.
        /// </returns>
        private static IPAddress GetDefaultIp()
        {
            IPAddress machineAddress = IPAddress.Loopback;

            // Find IPV4 Address
            foreach (IPAddress machineIP in Dns.GetHostEntry(Dns.GetHostName()).AddressList)
            {
                if (machineIP.AddressFamily == AddressFamily.InterNetwork)
                {
                    machineAddress = machineIP;
                    break;
                }
            }

            return machineAddress;
        }

        /// <summary>
        /// This function is used to retrieve the IP address collection from machine name.
        /// </summary>
        /// <param name="machineName">Name of the machine.</param>
        /// <returns>Collection of IP Addresses for the machine name</returns>
        private static Collection<IPAddress> GetIpFromName(string machineName)
        {
            Collection<IPAddress> machineAddresses = new Collection<IPAddress>();
            try
            {
                if (string.Compare(machineName, Constants.Localhost, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    machineAddresses.Add(GetDefaultIp());
                }
                else
                {
                    foreach (IPAddress machineIP in Dns.GetHostEntry(machineName).AddressList)
                    {
                        if (machineIP.AddressFamily == AddressFamily.InterNetwork)
                        {
                            machineAddresses.Add(machineIP);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.LogException(exception);
                throw new CustomException(Properties.Resources.RetrieveIPAddressFailure, exception, true);
            }

            return machineAddresses;
        }

        /// <summary>
        /// Sets the machine IP if the machine address is valid in following scenarios
        /// 1. WWT is running
        /// 2. It has version more than base version
        /// </summary>
        /// <param name="machineAddress">Machine Address</param>
        /// <param name="displayName">Machine display name</param>
        private void SetMachineIPIfValid(IPAddress machineAddress, string displayName)
        {
            if (machineAddress != null)
            {
                try
                {
                    // Checks if the machine is valid and latest WWT running on it.
                    if (WWTManager.IsValidMachine(machineAddress.ToString(), false))
                    {
                        MachineIP = machineAddress;
                        DisplayValue = displayName;
                    }
                }
                catch (CustomException ex)
                {
                    throw new CustomException(ex.HasCustomMessage ? ex.Message : Properties.Resources.DefaultErrorMessage, ex, true);
                }
            }
        }

        /// <summary>
        ///  Sets the machine IP if the machine address is valid in following scenarios
        /// 1. WWT is running
        /// 2. It has version more than base version
        /// </summary>
        /// <param name="machineAddresses">Collection of IP address for the name provided</param>
        /// <param name="machineName">Machine display name</param>
        private void SetMachineIPIfValid(Collection<IPAddress> machineAddresses, string machineName)
        {
            if (machineAddresses != null && machineAddresses.Count > 0)
            {
                int count = 0;
                foreach (IPAddress machineAddress in machineAddresses)
                {
                    try
                    {
                        count++;

                        // Checks if the machine is valid and latest WWT running on it.
                        if (WWTManager.IsValidMachine(machineAddress.ToString(), false))
                        {
                            MachineIP = machineAddress;
                            DisplayValue = machineName;
                            break;
                        }
                    }
                    catch (CustomException ex)
                    {
                        switch (ex.ErrorCode)
                        {
                            case ErrorCodes.Code100001:
                                // If the error code is for WWT not running and is the last IP address in the loop 
                                // then throw exception else continue with the looping
                                if (count == machineAddresses.Count)
                                {
                                    throw new CustomException(ex.HasCustomMessage ? ex.Message : Properties.Resources.DefaultErrorMessage, ex, true);
                                }
                                else
                                {
                                    continue;
                                }
                            case ErrorCodes.Code100003:
                            case ErrorCodes.Code100005:  
                            case ErrorCodes.Code100002:  
                                // If WWT is running on the machine but it not latest version then throw the exception
                                {
                                    throw new CustomException(ex.HasCustomMessage ? ex.Message : Properties.Resources.DefaultErrorMessage, ex, true);
                                }
                            default:
                                break;
                        }
                    }
                }
            }
        }
    }
}
