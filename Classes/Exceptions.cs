/*  Copyright (C) 2014 NAVERTICA a.s. http://www.navertica.com 

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
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using Microsoft.SharePoint;
using Navertica.SharePoint.Extensions;

// ReSharper disable once CheckNamespace
namespace Navertica.SharePoint
{
    public class ArgumentValidationException : Exception
    {
        public ArgumentValidationException(string argument, string message) : base("Parameter " + argument + " : " + message) {}
    }

    public class DirectoryEntryException : Exception
    {
        public DirectoryEntryException() {}
        public DirectoryEntryException(DirectoryEntry de) : base(GetMessage(de)) {}

        private static string GetMessage(DirectoryEntry de)
        {
            try
            {
                // ReSharper disable UnusedVariable
                Guid guid = de.Guid; //cause an exception
                // ReSharper restore UnusedVariable
            }
            catch (Exception exc)
            {
                return "Failed to Load DirectoryEntry : " + exc.Message;
            }
            return "";
        }
    }

    //http://www.codeproject.com/Articles/9538/Exception-Handling-Best-Practices-in-NET
    public class SPFieldNotFoundException : Exception
    {
        public SPFieldNotFoundException() {}
        public SPFieldNotFoundException(SPList list, string fieldIntName) : base(GetMessage(list, fieldIntName, "")) {}
        public SPFieldNotFoundException(SPList list, string fieldIntName, string additionalMessage) : base(GetMessage(list, fieldIntName, additionalMessage)) {}
        public SPFieldNotFoundException(SPList list, Guid guid) : base(GetMessage(list, guid)) {}
        public SPFieldNotFoundException(SPList list, IEnumerable<string> fieldIntNames) : base(GetMessage(list, fieldIntNames, "")) {}
        public SPFieldNotFoundException(SPList list, IEnumerable<string> fieldIntNames, string additionalMessage) : base(GetMessage(list, fieldIntNames, additionalMessage)) {}
        public SPFieldNotFoundException(SPList list, IEnumerable<Guid> guids) : base(GetMessage(list, guids)) {}

        private static string GetMessage(SPList list, string fieldIntName, string additionalMessage)
        {
            return GetMessage(list, new[] { fieldIntName }, additionalMessage);
        }

        private static string GetMessage(SPList list, IEnumerable<string> fieldIntNames, string additionalMessage)
        {
            List<string> notPres = fieldIntNames.Where(fieldIntName => !list.ContainsFieldIntName(fieldIntName)).Select(fieldIntName => "'" + fieldIntName + "'").ToList();
            return list.AbsoluteUrl() + " does not contains fields with internal names: " + notPres.JoinStrings(", ") + ( !string.IsNullOrEmpty(additionalMessage) ? "AdditionalMessage:" + additionalMessage + "\n" : "" );
        }

        private static string GetMessage(SPList list, Guid guid)
        {
            return GetMessage(list, new[] { guid });
        }

        private static string GetMessage(SPList list, IEnumerable<Guid> guids)
        {
            List<string> notPres = guids.Where(g => !list.ContainsFieldGuid(g)).Select(g => "'" + g + "'").ToList();
            return list.AbsoluteUrl() + " does not contains fields with guids: " + notPres.JoinStrings(", ");
        }
    }

    public class SPSecurableObjectAccesDeniedException : Exception
    {
        public SPSecurableObjectAccesDeniedException() {}
        public SPSecurableObjectAccesDeniedException(SPSecurableObject obj) : base(GetMessage(obj)) {}

        private static string GetMessage(SPSecurableObject obj)
        {
            string result = "";

            switch (obj.GetType().Name)
            {
                case "SPListItem":
                    result = "for SPListItem - " + ( (SPListItem) obj ).FormUrlDisplay();
                    break;
                case "SPList":
                    result = "for SPList - " + ( (SPList) obj ).DefaultViewUrl;
                    break;
                case "SPWeb":
                    result = "for SPWeb - " + ( (SPWeb) obj ).Url;
                    break;

            }

            return result;
        }
    }

    public class SPListAccesDeniedException : Exception
    {
        public SPListAccesDeniedException() {}
        public SPListAccesDeniedException(Guid guid, SPWeb web) : base("for list " + web.RunElevated(elevWeb => elevWeb.Lists[guid].DefaultViewUrl) + " with Guid '" + guid + "' at web " + web.Url) {}
        public SPListAccesDeniedException(string identification, SPWeb web) : base("for list " + web.RunElevated(elevWeb => elevWeb.OpenList(identification).DefaultViewUrl) + " with Guid '" + web.RunElevated(elevWeb => elevWeb.OpenList(identification).ID) + "' at web " + web.Url) {}
    }

    public class SPListItemNotFoundException : Exception
    {
        public SPListItemNotFoundException() {}
        public SPListItemNotFoundException(int id, SPList list) : base("SPListItem with Id '" + id + "' was not found at list " + list.AbsoluteUrl() + " or user have not permissions") {}
    }

    public class SPListNotFoundException : Exception
    {
        public SPListNotFoundException() {}
        public SPListNotFoundException(Guid guid, SPWeb web) : base("SPList with Guid '" + guid + "' was not found at web " + web.Url) {}
        public SPListNotFoundException(string identification, SPWeb web) : base("SPList with '" + identification + "' was not found at web " + web.Url) {}
    }

    public class SPWebAccesDeniedException : Exception
    {
        public SPWebAccesDeniedException() {}
        public SPWebAccesDeniedException(Guid guid, SPSite site) : base("Acces Denied to SPWeb with Guid '" + guid + "' at site " + site.Url) {}
        public SPWebAccesDeniedException(string identification, SPSite site) : base("Acces Denied to SPWeb '" + identification + "' at site " + site.Url) {}
    }

    public class SPWebNotFoundException : Exception
    {
        public SPWebNotFoundException() {}
        public SPWebNotFoundException(Guid guid, SPSite site) : base("SPWeb with Guid '" + guid + "' was not found at site " + site.Url) {}
        public SPWebNotFoundException(string identification, SPSite site) : base("SPWeb with '" + identification + "' was not found at site " + site.Url) {}
    }

    public class TerminateException : Exception {}
}