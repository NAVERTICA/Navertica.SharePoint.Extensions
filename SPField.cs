﻿/*  Copyright (C) 2014 NAVERTICA a.s. http://www.navertica.com 

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
using System.Linq;
using System.Reflection;
using Microsoft.SharePoint;

namespace Navertica.SharePoint.Extensions
{
    public static class SPFieldExtensions
    {
        /// <summary>
        /// Checks whether the list contains all the fields with internal names passed in intFieldNames
        /// </summary>
        /// <param name="fieldCollection"></param>
        /// <param name="guids">guids the list should contain</param>
        /// <returns></returns>
        public static bool ContainsFieldGuid(this SPFieldCollection fieldCollection, IEnumerable<Guid> guids)
        {
            if (fieldCollection == null) throw new ArgumentNullException("fieldCollection");
            if (guids == null) throw new ArgumentNullException("guids");

            foreach (Guid guid in guids)
            {
                try
                {
                    // ReSharper disable once UnusedVariable
                    var tmp = fieldCollection[guid];
                }
                catch
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Checks whether the list contains a field with given internal name
        /// </summary>
        /// <param name="fieldCollection"></param>
        /// <param name="guid"></param>
        /// <returns></returns>
        public static bool ContainsFieldGuid(this SPFieldCollection fieldCollection, Guid guid)
        {
            if (fieldCollection == null) throw new ArgumentNullException("fieldCollection");
            if (guid == null) throw new ArgumentNullException("guid");

            return ContainsFieldGuid(fieldCollection, new[] { guid });
        }

        /// <summary>
        /// Checks whether the field collection contains all given fields' internal names
        /// </summary>
        /// <param name="fieldCollection"></param>
        /// <param name="intFieldNames"></param>
        /// <returns></returns>
        public static bool ContainsFieldIntName(this SPFieldCollection fieldCollection, IEnumerable<string> intFieldNames)
        {
            if (fieldCollection == null) throw new ArgumentNullException("fieldCollection");
            if (intFieldNames == null) throw new ArgumentNullException("intFieldNames");

            foreach (string fieldname in intFieldNames)
            {
                try
                {
                    fieldCollection.GetFieldByInternalName(fieldname);
                }
                catch
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Checks whether the field collection contains a field with given internal name
        /// </summary>
        /// <param name="fieldCollection"></param>
        /// <param name="intFieldName"></param>
        /// <returns></returns>
        public static bool ContainsFieldIntName(this SPFieldCollection fieldCollection, string intFieldName)
        {
            if (fieldCollection == null) throw new ArgumentNullException("fieldCollection");
            if (intFieldName == null) throw new ArgumentNullException("intFieldName");

            return ContainsFieldIntName(fieldCollection, new[] { intFieldName });
        }

        /// <summary>
        /// For a given non-empty lookup field, opens the looked-up item (using itemId on the field's ParentList) and gets its WebListItemId        
        /// </summary>
        /// <param name="itemId"></param>
        /// <param name="lookupField"></param>
        /// <returns>WebListItemId</returns>
        public static WebListItemId GetItemFromLookup(this SPFieldLookup lookupField, int itemId)
        {
            if (lookupField == null) throw new ArgumentNullException("lookupField");

            WebListItemId result = new WebListItemId();

            try
            {
                using (SPWeb lookupWeb = lookupField.ParentList.ParentWeb.Site.OpenW(lookupField.LookupWebId, true))
                {
                    try
                    {
                        SPList lookupList = lookupWeb.OpenList(lookupField.LookupList, true);
                        SPListItem newItem = lookupList.GetItemById(itemId);
                        result = new WebListItemId(newItem);
                    }
                    catch (Exception listExc)
                    {
                        result.InvalidMessage = listExc.ToString();
                    }
                }
            }
            catch (Exception webExc)
            {
                result.InvalidMessage = webExc.ToString();
            }

            return result;
        }

        /// <summary>
        /// For constructing SPFieldLookupValue using only the field and item ID
        /// </summary>
        /// <param name="lookupField"></param>
        /// <param name="id"></param>
        /// <returns></returns>
        public static SPFieldLookupValue GetLookupValueForId(this SPFieldLookup lookupField, int id)
        {
            if (lookupField == null) throw new ArgumentNullException("lookupField");
            if (id < 1) throw new ArgumentException("id < 1");

            SPList lookupList = lookupField.ParentList.ParentWeb.OpenList(lookupField.LookupList, true);
            SPListItem item = lookupList.GetItemById(id);
            Guid lookupFieldId;
            try
            {
                lookupFieldId = new Guid(lookupField.LookupField);
            }
            catch
            {
                return new SPFieldLookupValue(id, ( item[lookupField.LookupField] ?? "" ).ToString());
            }

            return new SPFieldLookupValue(id, ( item[lookupFieldId] ?? "" ).ToString());
        }

        /// <summary>
        /// Returns a collection of internal names of all non-hidden and non-underscored fields
        /// </summary>
        /// <param name="fieldCollection"></param>
        /// <returns></returns>
        public static IEnumerable<string> InternalFieldNames(this SPFieldCollection fieldCollection)
        {
            if (fieldCollection == null) throw new ArgumentNullException("fieldCollection");

            return fieldCollection.Cast<SPField>().Where(fld => !fld.Hidden && fld.InternalName[0] != '_').Select(fld => fld.InternalName);
        }

        /// <summary>
        /// Checks whether field.TypeAsString contains the word "lookup", so that it works also for fieldtypes like FilteredLookup
        /// </summary>
        /// <param name="field"></param>
        /// <returns></returns>
        public static bool IsLookup(this SPField field)
        {
            if (field == null) throw new ArgumentNullException("field");

            string fieldType = field.TypeAsString.ToLowerInvariant();
            return fieldType.Contains("lookup") && fieldType != "extendedlookup"; // NAVERTICA ExtendedLookup is needed for backwards compatibility, and is not a lookup
        }

        /// <summary>
        /// Sets 'hidden' attr even for fields which don't allow it. 
        /// http://go4answers.webhost4life.com/Example/programatically-change-hidden-value-41576.aspx
        /// 
        /// </summary>
        /// <param name="field"></param>
        /// <param name="hidden"></param>
        /// <param name="inAllContenteTypes"></param>
        /// <returns>True if change was successful</returns>
        public static bool SetHidden(this SPField field, bool hidden, bool inAllContenteTypes = true)
        {
            if (field == null) throw new ArgumentNullException("field");

            try
            {
                if (field.Hidden != hidden)
                {
                    Type type = field.GetType();
                    MethodInfo mi = type.GetMethod("SetFieldBoolValue", BindingFlags.NonPublic | BindingFlags.Instance);
                    mi.Invoke(field, new object[] { "CanToggleHidden", true });
                    field.Hidden = hidden;
                    field.Update();
                }

                if (inAllContenteTypes)
                {
                    foreach (SPContentType ct in field.ParentList.ContentTypes)
                    {
                        if (ct.Fields.ContainsFieldIntName(field.InternalName)) //nemusi byt nutne ve vsech contenttypech
                        {
                            try
                            {
                                SPField fld = ct.Fields.GetFieldByInternalName(field.InternalName);

                                if (ct.FieldLinks[fld.Id].Hidden != hidden)
                                {
                                    ct.FieldLinks[fld.Id].Hidden = hidden;
                                    ct.Update();
                                }
                            }
                            catch (Exception) {}
                        }
                    }
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static string GetLookupValue(this SPFieldLookupValue lookup)
        {
            if (lookup == null) return null;
            return lookup.LookupValue;
        }

        public static string[] GetLookupValues(this SPFieldLookupValueCollection lookups)
        {
            if (lookups == null) return null;
            return lookups.Select(v => v.LookupValue).ToArray();
        }

        public static int GetLookupIndex(this SPFieldLookupValue lookup)
        {
            if (lookup == null) return -1;
            return lookup.LookupId;
        }

        public static int[] GetLookupIndexes(this SPFieldLookupValueCollection lookups)
        {
            if (lookups == null) return null;
            return lookups.Select(v => v.LookupId).ToArray();
        }
    }
}