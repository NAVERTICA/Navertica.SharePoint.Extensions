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
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Workflow;

namespace Navertica.SharePoint.Extensions
{
    public static partial class SPListItemExtensions
    {
        private static readonly Dictionary<string, string> FieldnameTranslation = new Dictionary<string, string>
        {
            { "vti_title", "Title" },
            { "vti_modifiedby", "Editor" },
            { "vti_author", "Author" },
            { "vti_nexttolasttimemodified", "Modified" },
            { "vti_timecreated", "Created" }
        };

        #region GET

        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
        /// <param name="fieldOrIdent"></param>
        /// <param name="singleValueOnly">in case of multi-value fields (users, lookups) set this to true if you want only the first value</param>
        /// <returns></returns>
        public static object Get(this SPListItem item, object fieldOrIdent, bool singleValueOnly = false)
        {
            if (item == null) throw new ArgumentNullException("item");
            if (fieldOrIdent == null) throw new ArgumentNullException("fieldOrIdent");
            SPField fld;

            if (fieldOrIdent is SPField) fld = fieldOrIdent as SPField;
            else fld = item.ParentList.OpenField(fieldOrIdent.ToString());

            return Get(item[fld.InternalName], fld, singleValueOnly);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="itemVersion"></param>
        /// <param name="fieldOrIdent"></param>
        /// <param name="singleValueOnly">in case of multi-value fields (users, lookups) set this to true if you want only the first value</param>
        /// <returns></returns>
        public static object Get(this SPListItemVersion itemVersion, object fieldOrIdent, bool singleValueOnly = false)
        {
            if (itemVersion == null) throw new ArgumentNullException("itemVersion");
            if (fieldOrIdent == null) throw new ArgumentNullException("fieldOrIdent");

            SPField fld;

            if (fieldOrIdent is SPField) fld = fieldOrIdent as SPField;
            else fld = itemVersion.ListItem.ParentList.OpenField(fieldOrIdent.ToString());

            return Get(itemVersion[fld.InternalName], fld, singleValueOnly);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="itemEventProperties"></param>
        /// <param name="fieldOrIdent"></param>
        /// <param name="singleValueOnly">in case of multi-value fields (users, lookups) set this to true if you want only the first value</param>
        /// <returns></returns>
        public static object Get(this SPItemEventProperties itemEventProperties, object fieldOrIdent, bool singleValueOnly = false)
        {
            if (itemEventProperties == null) throw new ArgumentNullException("itemEventProperties");
            if (fieldOrIdent == null) throw new ArgumentNullException("fieldOrIdent");

            SPField fld;

            if (fieldOrIdent is SPField) fld = fieldOrIdent as SPField;
            else fld = itemEventProperties.List.OpenField(fieldOrIdent.ToString());

            string fieldInternalName = fld.InternalName;

            // no AfterProperties in these events, use SPListItem.Get directly
            if (itemEventProperties.EventType == SPEventReceiverType.ItemAdded || itemEventProperties.EventType == SPEventReceiverType.ItemUpdated)
            {
                return Get(itemEventProperties.ListItem[fld.InternalName], fld, singleValueOnly);
            }

            Dictionary<string, object> changedProperties = itemEventProperties.AfterProperties.Cast<DictionaryEntry>().ToDictionary(pair => pair.Key.ToString(), pair => pair.Value);

            if (FieldnameTranslation.Values.Contains(fieldInternalName))
            {
                Dictionary<string, string> reverseTranslation = FieldnameTranslation.ToDictionary(x => x.Value, x => x.Key);
                fieldInternalName = reverseTranslation[fieldInternalName];
            }

            if (changedProperties.ContainsKey(fieldInternalName))
            {
                return Get(itemEventProperties.AfterProperties[fieldInternalName], fld, singleValueOnly);
            }

            if (itemEventProperties.ListItem == null) return null; // null when ItemAdding 

            return Get(itemEventProperties.ListItem[fld.InternalName], fld, singleValueOnly);
        }

        /// <summary>
        /// Get values for multiple fields from item in a dictionary. In case of exception when reading field, the exception will be the resulting value.
        /// </summary>
        /// <param name="item"></param>
        /// <param name="fieldIds"></param>
        /// <param name="singleValueOnly">in case of multi-value fields (users, lookups) set this to true if you want only the first value</param>
        /// <returns></returns>
        public static DictionaryNVR Get(this SPListItem item, IEnumerable<string> fieldIds, bool singleValueOnly = false)
        {
            if (item == null) throw new ArgumentNullException("item");
            if (fieldIds == null) throw new ArgumentNullException("fieldIds");

            DictionaryNVR results = new DictionaryNVR();

            foreach (string fieldOrIdent in fieldIds)
            {
                SPField fld = item.ParentList.OpenField(fieldOrIdent);
                results[fld.InternalName] = item.Get(fld, singleValueOnly);
            }

            return results;
        }

        /// <summary>
        /// Get values for multiple fields from item's version in a dictionary. In case of exception when reading field, the exception will be the resulting value.
        /// </summary>
        /// <param name="itemVersion"></param>
        /// <param name="fieldIds"></param>
        /// <param name="singleValueOnly">in case of multi-value fields (users, lookups) set this to true if you want only the first value</param>
        /// <returns></returns>
        public static DictionaryNVR Get(this SPListItemVersion itemVersion, IEnumerable<string> fieldIds, bool singleValueOnly = false)
        {
            if (itemVersion == null) throw new ArgumentNullException("itemVersion");
            if (fieldIds == null) throw new ArgumentNullException("fieldIds");

            DictionaryNVR results = new DictionaryNVR();

            foreach (string fieldOrIdent in fieldIds)
            {
                SPField fld = itemVersion.ListItem.ParentList.OpenField(fieldOrIdent);
                results[fld.InternalName] = itemVersion.Get(fld, singleValueOnly);
            }

            return results;
        }

        /// <summary>
        /// Get values for multiple fields from event properties in a dictionary. In case of exception when reading field, the exception will be the resulting value.
        /// </summary>
        /// <param name="itemEventProperties"></param>
        /// <param name="fieldIds"></param>
        /// <param name="singleValueOnly">in case of multi-value fields (users, lookups) set this to true if you want only the first value</param>
        /// <returns></returns>
        public static DictionaryNVR Get(this SPItemEventProperties itemEventProperties, IEnumerable<string> fieldIds, bool singleValueOnly = false)
        {
            if (itemEventProperties == null) throw new ArgumentNullException("itemEventProperties");
            if (fieldIds == null) throw new ArgumentNullException("fieldIds");

            DictionaryNVR results = new DictionaryNVR();

            foreach (string fieldOrIdent in fieldIds)
            {
                SPField fld = itemEventProperties.List.OpenField(fieldOrIdent);
                results[fld.InternalName] = itemEventProperties.Get(fld, singleValueOnly);
            }

            return results;
        }

        /// <summary>
        /// Workhorse method, in case of exception the exception will be returned as the result
        /// </summary>
        /// <param name="value"></param>
        /// <param name="field"></param>
        /// <param name="singleValueOnly">in case of multi-value fields (users, lookups) set this to true if you want only the first value</param>
        /// <returns></returns>
        private static object Get(object value, SPField field, bool singleValueOnly = false)
        {
            if (value == null) 
            {
                if (field.GetType() == typeof (SPFieldBoolean)) return false;
                return null;
            }

            if (field.GetType() == typeof (SPField) && field.Type == SPFieldType.Counter) return value.ToInt();

            if (field.InternalName == "ContentTypeId")
            {
                SPContentTypeId ctid = new SPContentTypeId(value.ToString());
                return ctid;
            }

            if (field.GetType() == typeof (SPFieldText) || field.GetType() == typeof (SPFieldMultiLineText))
            {
                return value.ToString();
            }

            try
            {
                //Get specific type of field
                Type t = field.GetType().UnderlyingSystemType;
                List<object> pars = new List<object>();

                //Try to load Get method from Custom Field
                MethodInfo method = field.GetType().GetMethod("Get");
                if (method != null)
                {
                    try
                    {
                        pars.Add(value);
                        if (singleValueOnly) pars.Add(true);

                        object r = method.Invoke(field, pars.ToArray());
                        return r;
                    }
                    catch (Exception exc)
                    {
                        return exc;
                    }
                }

                //Try to load Get method from extensions
                method = t.GetExtensionMethod("Get");
                if (method != null)
                {
                    pars.Add(field);
                    pars.Add(value);
                    if (singleValueOnly) pars.Add(true);

                    object r = method.Invoke(null, pars.ToArray());
                    return r;
                }

                //Field doesnt have specified Get method so return the pure value)
                
                return value;
            }
            catch (Exception exc)
            {
                return exc;
            }
        }

        #endregion

        #region SET

        public static void Set(this SPListItem item, string fieldOrIdent, object value)
        {
            if (item == null) throw new ArgumentNullException("item");
            if (fieldOrIdent == null) throw new ArgumentNullException("fieldOrIdent");

            Set((object) item, new DictionaryNVR { { fieldOrIdent, value } });
        }

        public static void Set(this SPListItem item, IDictionary<string, object> dict)
        {
            if (item == null) throw new ArgumentNullException("item");
            if (dict == null) throw new ArgumentNullException("dict");

            Set((object) item, dict);
        }

        public static void Set(this SPItemEventProperties itemEventProperties, string fieldOrIdent, object value)
        {
            if (itemEventProperties == null) throw new ArgumentNullException("itemEventProperties");
            if (fieldOrIdent == null) throw new ArgumentNullException("fieldOrIdent");

            Set((object) itemEventProperties, new DictionaryNVR { { fieldOrIdent, value } });
        }

        public static void Set(this SPItemEventProperties itemEventProperties, IDictionary<string, object> dict)
        {
            if (itemEventProperties == null) throw new ArgumentNullException("itemEventProperties");
            if (dict == null) throw new ArgumentNullException("dict");

            Set((object) itemEventProperties, dict);
        }

        private static void Set(object spItem, IDictionary<string, object> dict)
        {
            if (dict == null) return;
            if (dict.Count == 0) return;

            SPListItem item = spItem.GetType() == typeof (SPListItem) ? ( (SPListItem) spItem ) : null;
            SPItemEventProperties itemEventProperties = spItem is SPItemEventProperties ? ( (SPItemEventProperties) spItem ) : null;

            // ReSharper disable once PossibleNullReferenceException
            SPList list = item != null ? item.ParentList : itemEventProperties.List;

            if (!list.ContainsFieldIntName(dict.Keys))
            {
                throw new SPFieldNotFoundException(list, dict.Keys);
            }

            foreach (var kvp in dict)
            {
                SPField field = list.OpenField(kvp.Key);

                object valueToSave = kvp.Value;

                if (valueToSave != null)
                {
                    if (field.GetType() == typeof (SPFieldBoolean))
                    {
                        valueToSave = kvp.Value.ToBool();
                    }

                    if (field.GetType() == typeof (SPFieldDateTime))
                    {
                        DateTime? date = (DateTime?) ( (SPFieldDateTime) field ).Get(valueToSave);

                        if (date.HasValue)
                        {
                            valueToSave = (DateTime) date;

                            if (itemEventProperties != null)
                            {
                                valueToSave = ( (DateTime) valueToSave ).ToStringISO();
                            }
                        }
                        else
                        {
                            valueToSave = null;
                        }
                    }

                    if (field.GetType() == typeof (SPFieldUrl) && valueToSave != null)
                    {
                        if (valueToSave.GetType() != typeof (SPFieldUrlValue))
                        {
                            SPFieldUrlValue url = new SPFieldUrlValue((string) valueToSave);
                            valueToSave = url;
                        }
                    }

                    if (field.GetType() == typeof (SPFieldLookup))
                    {
                        valueToSave = SetSPFieldLookup((SPFieldLookup) field, valueToSave, item);
                    }

                    if (field.GetType() == typeof (SPFieldUser))
                    {
                        valueToSave = SetSPFieldUser((SPFieldUser) field, valueToSave, item);
                    }

                    if (field.GetType() == typeof (SPFieldGuid)) ; // unable to set this type of field                    
                    if (field.GetType() == typeof (SPFieldModStat)) ;
                    if (field.GetType() == typeof (SPFieldWorkflowStatus)) ;
                    if (field.GetType() == typeof (SPFieldCalculated)) ; // nothing to set
                }

                //STORE THE VALUE
                if (item != null)
                {
                    item[field.InternalName] = valueToSave;
                }
                else
                {
                    itemEventProperties.AfterProperties[field.InternalName] = valueToSave;
                }
            }
        }

        private static object SetSPFieldLookup(SPFieldLookup lookup, object valueToSave, SPListItem item)
        {
            if (valueToSave.ToString() == "") return null;

            string lookupString = valueToSave.ToString();

            if (lookup.AllowMultipleValues)
            {
                SPFieldLookupValueCollection lvc = new SPFieldLookupValueCollection();
                if (valueToSave.GetType() == typeof (SPFieldLookupValueCollection))
                {
                    lvc = (SPFieldLookupValueCollection) valueToSave;
                }
                else
                {
                    if (lookupString.Contains(";#"))
                    {
                        lvc = new SPFieldLookupValueCollection(lookupString);
                    }
                    else
                    {
                        if (valueToSave.GetType().GetInterface("IEnumerable") != null)
                        {
                            List<int> ids = new List<int>();
                            foreach (object o in (IEnumerable) valueToSave)
                            {
                                ids.Add(o.ToInt());
                            }
                            lvc = new SPFieldLookupValueCollection(ids.JoinStrings(";#;#"));
                        }
                    }
                }

                if (item != null)
                {
                    valueToSave = lvc;
                }
                else
                {
                    valueToSave = lvc.ToString();
                }
            }
            else //single lookup
            {
                if (valueToSave.GetType() == typeof (SPFieldLookupValue))
                {
                    valueToSave = ( (SPFieldLookupValue) valueToSave ).LookupId;
                }
                else
                {
                    if (lookupString.Contains(";#")) //lookup string
                    {
                        valueToSave = lookupString.GetLookupIndex();
                    }
                    else
                    {
                        try
                        {
                            valueToSave = int.Parse(valueToSave.ToString());
                        }
                        catch
                        {
                            valueToSave = null;
                        }
                    }
                }
            }

            return valueToSave;
        }

        private static object SetSPFieldUser(SPFieldUser userField, object valueToSave, SPListItem item)
        {
            List<SPPrincipal> principals = userField.ParentList.ParentWeb.GetSPPrincipals(valueToSave);

            if (principals.Count > 0)
            {
                if (item != null)
                {
                    valueToSave = principals.GetSPFieldUserValueCollection();
                }
                else
                {
                    //http://naimmurati.wordpress.com/2014/07/25/user-field-in-event-receivers-when-using-claims-based-authentication-and-classic-mode-authentication/

                    if (principals.Any(p => p.LoginName.Contains('#')))
                    {
                        List<string> users = new List<string>();
                        foreach (SPPrincipal principal in principals)
                        {
                            users.Add(principal.ID + ";#" + principal.LoginName);
                        }
                        valueToSave = users.JoinStrings(";#");
                    }
                    else
                    {
                        foreach (SPPrincipal principal in principals)
                        {
                            valueToSave += principal.ID + ";#" + principal.LoginName.Split('\\')[1] + ";";
                        }
                    }
                }
            }
            else
            {
                valueToSave = null;
            }

            return valueToSave;
        }

        #endregion
    }
}