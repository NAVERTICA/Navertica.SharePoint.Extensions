/*  Copyright (C) 2015 NAVERTICA a.s. http://www.navertica.com 

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
using System.Globalization;
using System.Linq;
using System.Reflection;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Workflow;

namespace Navertica.SharePoint.Extensions
{
    public static class SPFieldGetExtensions
    {
        #region GET FIELD EXTENSIONS

        public static bool Get(this SPFieldBoolean fld, object value)
        {
            if (value == null) return false;

            // TODO - kdyz je pole typu bool a je nastavene na True, v AfterProperties je pry hodnota -1, a to potrebujeme osetrit
            // viz Remarks - http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.spitemeventproperties.afterproperties.aspx#Remarks
            if (value.ToString() == "-1") return true;

            return value.ToBool();
        }

        public static object Get(this SPFieldCalculated fld, object value)
        {
            //TODO - dalsi datove typy
            string[] kvp = value.ToString().Split(";#");
            switch (kvp[0])
            {
                case "boolean":
                    return kvp[1].ToBool();
                case "datetime":
                    return GetDateTime(fld, kvp[1]);
                case "float":
                    float res;
                    float.TryParse(kvp[1], NumberStyles.Float, CultureInfo.InvariantCulture.NumberFormat, out res);
                    return res;
                case "string":
                    return kvp[1];
                default:
                    return kvp[1];
            }
        }

        public static object Get(this SPFieldComputed fld, object value)
        {
            //TODO - what todo with computed items?Maybe some will be specific
            return value;
        }

        public static double Get(this SPFieldCurrency fld, object value)
        {
            return value.ToDouble();
        }

        public static object Get(this SPFieldDateTime fld, object value)
        {
            return GetDateTime(fld, value);
        }

        private static DateTime? GetDateTime(this SPField fld, object value)
        {
            if (value is DateTime) return (DateTime?)value;

            DateTime res;
            DateTime.TryParse(value.ToString(), out res);

            if (res.Year != 1) return res;

            // parse unsuccessful
            string dateString = value.ToString();

            DateTime? date = null;
            try
            {
                date = DateTime.Parse(dateString);
            }
            catch
            {
                DateTime dateAttempt;
                foreach (int lcid in fld.ParentList.ParentWeb.RegionalSettings.InstalledLanguages)
                {
                    CultureInfo info = new CultureInfo(lcid);
                    if (DateTime.TryParse(dateString, info, DateTimeStyles.None, out dateAttempt))
                    {
                        date = dateAttempt;
                        break;
                    }
                }
            }
           
            return date;
        }

        public static object Get(this SPFieldGuid fld, object value)
        {
            if (value == null) return Guid.Empty;

            //fld van contains SPFieldUrlValue?????
            if (value is SPFieldUrlValue) return value;

            return new Guid(value.ToString());
        }

        //Cant use optional parameter, due to reflection calling method
        public static object Get(this SPFieldLookup fld, object value)
        {
            return Get(fld, value, false);
        }

        /// <summary>
        /// Private method for use in Get
        /// </summary>
        /// <param name="fld"></param>
        /// <param name="value"></param>
        /// <param name="singleValueOnly">in case of multi-value fields (users, lookups) set this to true if you want only the first value</param>
        /// <returns></returns>
        public static object Get(this SPFieldLookup fld, object value, bool singleValueOnly)
        {
            if (value is SPFieldLookupValue) return value;

            var collection = value as SPFieldLookupValueCollection;
            if (collection != null)
            {
                if (collection.Count == 0) return null;

                return singleValueOnly ? (object)collection[0] : collection;
            }

            string strValue = value.ToString();

            if (strValue == "") return null;

            if (fld.AllowMultipleValues)
            {
                collection = new SPFieldLookupValueCollection(strValue);
                return singleValueOnly ? (object)collection[0] : collection;
            }

            return new SPFieldLookupValue(strValue);
        }

        public static SPModerationStatusType? Get(this SPFieldModStat fld, object value)
        {
            if (value == null) return null; //return what??

            int intValue = -1;

            if (value is string)
            {
                string stringValue = value.ToString();

                if (stringValue == "Approved") intValue = 0;
                if (stringValue == "Denied") intValue = 1;
                if (stringValue == "Pending") intValue = 2;
                if (stringValue == "Draft") intValue = 3;
                if (stringValue == "Scheduled") intValue = 4;
            }

            if (intValue < 0)
            {
                intValue = value.ToInt();
            }

            switch (intValue)
            {
                case 0:
                    return SPModerationStatusType.Approved;
                case 1:
                    return SPModerationStatusType.Denied;
                case 2:
                    return SPModerationStatusType.Pending;
                case 3:
                    return SPModerationStatusType.Draft;
                case 4:
                    return SPModerationStatusType.Scheduled;
                default:
                    return null;
            }
        }

        public static object Get(this SPFieldNumber fld, object value)
        {
            if (value == null) return 0;

            //TODO
            return null;
        }

        public static SPFieldUrlValue Get(this SPFieldUrl fld, object value)
        {
            if (value is SPFieldUrlValue) return (SPFieldUrlValue)value;

            return new SPFieldUrlValue(value.ToString());
        }

        //Cant use optional parameter, due to reflection calling method
        public static object Get(this SPFieldUser fld, object value)
        {
            return Get(fld, value, false);
        }

        /// <summary>
        /// Private method for use in Get
        /// </summary>
        /// <param name="fld"></param>
        /// <param name="value"></param>
        /// <param name="singleValueOnly">in case of multi-value fields (users, lookups) set this to true if you want only the first value</param>
        /// <returns></returns>
        public static object Get(this SPFieldUser fld, object value, bool singleValueOnly)
        {
            List<SPPrincipal> principals = fld.ParentList.ParentWeb.GetSPPrincipals(value);

            if (principals != null)
            {
                if (principals.Count == 0) return null;

                if (fld.SelectionMode == SPFieldUserSelectionMode.PeopleAndGroups)
                {
                    if (fld.AllowMultipleValues)
                    {
                        return singleValueOnly ? (object)principals[0] : principals;
                    }
                    return principals[0];
                }
                if (fld.SelectionMode == SPFieldUserSelectionMode.PeopleOnly)
                {
                    List<SPUser> users = fld.ParentList.ParentWeb.GetSPUsers(principals);
                    if (fld.AllowMultipleValues)
                    {
                        return singleValueOnly ? (object)users[0] : users;
                    }
                    return users[0];
                }
            }

            return null;
        }

        public static SPWorkflowStatus? Get(this SPFieldWorkflowStatus fld, object value)
        {
            if (value == null) return null; //return what??

            int intValue = -1;

            if (value is string)
            {
                string stringValue = value.ToString();

                if (stringValue == "NotStarted") intValue = 0;
                if (stringValue == "FailedOnStart") intValue = 1;
                if (stringValue == "InProgress") intValue = 2;
                if (stringValue == "ErrorOccurred") intValue = 3;
                if (stringValue == "StoppedByUser") intValue = 4;
                if (stringValue == "Completed") intValue = 5;
                if (stringValue == "FailedOnStartRetrying") intValue = 6;
                if (stringValue == "ErrorOccurredRetrying") intValue = 7;
                if (stringValue == "ViewQueryOverflow") intValue = 8;
            }

            if (intValue < 0)
            {
                intValue = value.ToInt();
            }

            switch (intValue)
            {
                //http://geekswithblogs.net/simonh/archive/2013/04/12/sharepoint-2010-workflow-status-values.aspx

                /*
            // - nejak tyto statusy chybi
            • Canceled = 15
            • Approved = 16
            • Rejected = 17
            */

                case 0:
                    return SPWorkflowStatus.NotStarted;
                case 1:
                    return SPWorkflowStatus.FailedOnStart;
                case 2:
                    return SPWorkflowStatus.InProgress;
                case 3:
                    return SPWorkflowStatus.ErrorOccurred;
                case 4:
                    return SPWorkflowStatus.StoppedByUser;
                case 5:
                    return SPWorkflowStatus.Completed;
                case 6:
                    return SPWorkflowStatus.FailedOnStartRetrying;
                case 7:
                    return SPWorkflowStatus.ErrorOccurredRetrying;
                case 8:
                    return SPWorkflowStatus.ViewQueryOverflow;

                default:
                    return SPWorkflowStatus.NotStarted;
            }
        }

        #endregion

    }
}
