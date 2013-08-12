// Visogram WinControl plugin
// Copyright (C) 2013  Yailo GbR - Michaelis & Romero
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.

// You should have received a copy of the GNU General Public License
// along with this program.  If not, see {http://www.gnu.org/licenses/}.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Visogram.Framework;
using Visogram.Groups;
using Visogram.Modules;
using System.Windows.Forms;
using System.Management;

namespace CustomModules.System
{
    [InitialState(InitialState.Enabled)]
    [Module(Group = typeof(ModuleGroupSystem))]
    public class WinControl : ModuleBase
    {
        public enum ControlFunction : int
        {
            Hibernate = 0xBE,
            Suspend = 0xEF,
            LogOff = 0x00,
            ForcedLogOff = 0x04,
            Shutdown = 0x01,
            ForcedShutdown = 0x05,
            Reboot = 0x02,
            ForcedReboot = 0x06,
            PowerOff = 0x08,
            ForcedPowerOff = 0x0C
        }

        /*** Define properties ***/
        private ControlFunction action;
        [Property]
        [Persistent]
        public ControlFunction Action
        {
            get { return action; }
            set { SetProperty("Action", ref action, value); }
        }

        /*** Define inputs ***/
        protected readonly SignalSink<object> inputTrigger;

        /*** Define outputs ***/

        /*** Module constructor ***/
        public WinControl()
        {
            inputTrigger = new SignalSink<object>(new EventHandler<SignalEventArgs<object>>(InputEventTrigger), "Trigger");
            AddInput(inputTrigger);

        }

        /*** Input event handlers ***/
        protected void InputEventTrigger(object sender, SignalEventArgs<object> e)
        {
            switch (Action)
            {
                //Hibernate and Standby can be performed easier...
                case ControlFunction.Hibernate:
                    Application.SetSuspendState(PowerState.Hibernate, true, false);
                    break;
                case ControlFunction.Suspend:
                    Application.SetSuspendState(PowerState.Suspend, true, false);
                    break;

                //For all other cases we need Win32 calls
                case ControlFunction.LogOff:
                case ControlFunction.ForcedLogOff:
                case ControlFunction.Shutdown:
                case ControlFunction.ForcedShutdown:
                case ControlFunction.Reboot:
                case ControlFunction.ForcedReboot:
                case ControlFunction.PowerOff:
                case ControlFunction.ForcedPowerOff:
                    OsManagementCall(Action);
                    break;
            }
        }

        private void OsManagementCall(ControlFunction flags)
        {
            ManagementBaseObject mboShutdown = null;
            ManagementClass mcWin32 = new ManagementClass("Win32_OperatingSystem");
            mcWin32.Get();

            //You can't shutdown without security privileges
            mcWin32.Scope.Options.EnablePrivileges = true;
            
            //Set function params
            ManagementBaseObject mboShutdownParams = mcWin32.GetMethodParameters("Win32Shutdown");
            mboShutdownParams["Flags"] = flags;
            mboShutdownParams["Reserved"] = 0;

            foreach (ManagementObject manObj in mcWin32.GetInstances())
            {
                mboShutdown = manObj.InvokeMethod("Win32Shutdown", mboShutdownParams, null);
            }

            if (mboShutdown == null)
                throw new Exception("Unable to perform Win32Shutdown call.");

            int res = Convert.ToInt32(mboShutdown["returnValue"]);
            if (res != 0)
                throw new Exception(String.Format("Win32Shutdown call returned error code {0}.", res));
        }
    }
}