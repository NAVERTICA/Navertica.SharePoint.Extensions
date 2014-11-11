/*  Copyright (C) 2014 NAVERTICA a.s.

    This program is free software; you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation; either version 2 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License along
    with this program; if not, write to the Free Software Foundation, Inc.,
    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.  */
using System;
using System.Reflection;
using Microsoft.SharePoint;

// ReSharper disable once CheckNamespace
namespace Navertica.SharePoint
{
    /// provides access to the internal Microsoft.SharePoint.SPEventManager class by using reflection
    /// sample usage:
    /// SPEventManagerWrapper.DisableEventFiring();

    /// SPList myList = SPContext.Current.Web.Lists["Shared Documents"];
    /// myList.Items[0].Update();

    /// SPEventManagerWrapper.EnableEventFiring();

    public static class SPEventManagerWrapper
    {
        private const string ClassName = "Microsoft.SharePoint.SPEventManager";
        private const string EventFiringSwitchName = "EventFiringDisabled";
        private static Type _eventManagerType;

        /// gets the status of event firing on the current thread
        public static bool EventFiringDisabled
        {
            get { return GetEventFiringSwitchValue(); }
        }

        private static Type EventManagerType
        {
            get
            {
                if (_eventManagerType == null) GetEventManagerType();

                return _eventManagerType;
            }
        }

        /// enables event firing on the current thread
        public static void EnableEventFiring()
        {
            SetEventFiringSwitch(false);
        }

        /// disables sharepoint event firing on the current thread
        public static void DisableEventFiring()
        {
            SetEventFiringSwitch(true);
        }

        /// sets the event firing switch on Microsoft.SharePoint.SPEventManager class using reflection
        private static void SetEventFiringSwitch(bool value)
        {
            PropertyInfo pi = EventManagerType.GetProperty(EventFiringSwitchName, BindingFlags.Static | BindingFlags.NonPublic);

            pi.SetValue(null, value, null);
        }

        private static bool GetEventFiringSwitchValue()
        {
            PropertyInfo pi = EventManagerType.GetProperty(EventFiringSwitchName, BindingFlags.Static | BindingFlags.NonPublic);

            object val = pi.GetValue(null, null);

            return (bool) val;
        }

        // ReSharper disable once UnusedMethodReturnValue.Local
        private static Type GetEventManagerType()
        {
            _eventManagerType = typeof (SPList).Assembly.GetType(ClassName, true);

            return _eventManagerType;
        }
    }
}