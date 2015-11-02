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
using System.Linq;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;

namespace Navertica.SharePoint.Extensions
{
    public static partial class SPListItemExtensions
    {

        public static DictionaryNVR GetItemDictionary(this SPListItem item, IEnumerable<string> fieldIntNames = null, bool expandLookups = false)
        {
            if (item == null) throw new ArgumentNullException("item");

            return GetItemDictionary(item, item.Fields, fieldIntNames, expandLookups ? new WebListItemDictionary() : null);
        }

        public static DictionaryNVR GetItemDictionary(this SPListItemVersion itemVersion, IEnumerable<string> fieldIntNames = null, bool expandLookups = false)
        {
            if (itemVersion == null) throw new ArgumentNullException("itemVersion");

            return GetItemDictionary(itemVersion, itemVersion.Fields, fieldIntNames, expandLookups ? new WebListItemDictionary() : null);
        }

        public static DictionaryNVR GetItemDictionary(this SPItemEventProperties properties, IEnumerable<string> fieldIntNames = null, bool expandLookups = false)
        {
            if (properties == null) throw new ArgumentNullException("properties");

            return GetItemDictionary(properties, properties.List.Fields, fieldIntNames, expandLookups ? new WebListItemDictionary() : null);
        }

        private static DictionaryNVR GetItemDictionary(object spItem, SPFieldCollection fields, IEnumerable<string> fieldIntNames = null, WebListItemDictionary loaded = null)
        {
            DictionaryNVR data = new DictionaryNVR();
            fieldIntNames = fieldIntNames != null ? fieldIntNames.ToList() : new List<string>();

            SPListItem item = spItem.GetType() == typeof (SPListItem) ? ( (SPListItem) spItem ) : null;
            SPListItemVersion itemVersion = spItem is SPListItemVersion ? ( (SPListItemVersion) spItem ) : null;
            SPItemEventProperties itemEventProperties = spItem is SPItemEventProperties ? ( (SPItemEventProperties) spItem ) : null;

            foreach (SPField field in fields)
            {
                object value = null;

                if (fieldIntNames.Any() && !field.InternalName.EqualAny(fieldIntNames)) continue;

                try
                {
                    #region ContentTypeId

                    if (field.InternalName == "ContentTypeId")
                    {
                        if (item != null)
                        {
                            data[field.InternalName] = ( (SPContentTypeId) item[field.InternalName] ).Parent.ToString();
                        }

                        if (itemEventProperties != null)
                        {
                            value = itemEventProperties.Get(field.InternalName);
                            if (value != null)
                            {
                                value = ( (SPContentTypeId) value ).Parent.ToString();
                            }
                            data[field.InternalName] = value;
                        }

                        if (itemVersion != null)
                        {
                            value = itemVersion.Get(field.InternalName);
                            if (value != null)
                            {
                                value = ( (SPContentTypeId) value ).Parent.ToString();
                            }
                            data[field.InternalName] = value;
                        }

                        continue;
                    }

                    #endregion

                    if (field.Hidden && item != null && !item.ContentType.ContainsFieldIntName(field.InternalName)) continue;
                    if (field.Hidden && itemVersion != null && !itemVersion.ListItem.ContentType.ContainsFieldIntName(field.InternalName)) continue;
                    if (field.Hidden && itemEventProperties != null && itemEventProperties.ListItem != null && !itemEventProperties.ListItem.ContentType.ContainsFieldIntName(field.InternalName)) continue;

                    if (item != null) value = item.Get(field.InternalName);
                    if (itemVersion != null) value = itemVersion.Get(field.InternalName);
                    if (itemEventProperties != null) value = itemEventProperties.Get(field.InternalName);

                    if (value is Exception)
                    {
                        if (!data.ContainsKey("_Exceptions")) data["_Exceptions"] = new DictionaryNVR();
                        ((DictionaryNVR)data["_Exceptions"])[field.InternalName] = value;
                        data[field.InternalName] = null;
                        continue;
                    }

                    if (value != null && field.GetType() == typeof (SPFieldUser))
                    {
                        List<SPPrincipal> principals;
                        if (value.GetType().GetInterface("IEnumerable") != null)
                        {
                            principals = (List<SPPrincipal>) value;
                        }
                        else
                        {
                            principals = new List<SPPrincipal> { (SPPrincipal) value };
                        }

                        List<SimpleSPPrincipal> simpleprincipals = principals.Select(principal => new SimpleSPPrincipal(principal)).ToList();

                        value = simpleprincipals.Count > 0 ? simpleprincipals : null;
                    }

                    if (value != null && field.GetType() == typeof (SPFieldLookup))
                    {
                        SPFieldLookup lookupField = (SPFieldLookup) field;

                        List<SPFieldLookupValue> lookups;

                        if (value.GetType().GetInterface("IEnumerable") != null)
                        {
                            lookups = ( (SPFieldLookupValueCollection) value ).ToList();
                        }
                        else
                        {
                            lookups = new List<SPFieldLookupValue> { ( (SPFieldLookupValue) value ) };
                        }

                        List<SimpleSPLookup> simpleLookups = new List<SimpleSPLookup>();

                        foreach (SPFieldLookupValue lookupValue in lookups)
                        {
                            if (lookupValue.LookupId == 0) continue;
                            simpleLookups.Add(new SimpleSPLookup(lookupField, lookupValue));
                        }
                        value = simpleLookups.Count > 0 ? simpleLookups : null;

                        if (loaded != null && simpleLookups != null)
                        {
                            foreach (var look in simpleLookups)
                            {
                                if (!loaded.Contains(look.WLI))
                                {
                                    look.WLI.ProcessItem(fields.Web.Site, delegate(SPListItem lookupItem)
                                    {
                                        WebListItemDictionary local = new WebListItemDictionary(loaded);
                                        local.Add(new WebListItemId(lookupItem));

                                        if (itemVersion == null || !lookupItem.ParentList.EnableVersioning)
                                        {
                                            value = GetItemDictionary(lookupItem, lookupItem.Fields, null, local);
                                        }
                                        else
                                        {
                                            SPListItemVersion versionToLoad = lookupItem.Versions.Cast<SPListItemVersion>().FirstOrDefault(v => v.Created < itemVersion.Created);

                                            if (versionToLoad == null)
                                            {
                                                versionToLoad = lookupItem.Versions[0];
                                            }

                                            value = GetItemDictionary(versionToLoad, lookupItem.Fields, null, local);
                                        }
                                        return null;
                                    });
                                }
                            }
                        }
                    }

                    #region Empty string to NULL

                    if (value is string)
                    {
                        string strValue = ( value ?? "" ).ToString();
                        if (strValue.Trim() == string.Empty || SPHttpUtility.HtmlDecode(strValue).Trim() == string.Empty)
                        {
                            value = null;
                        }
                    }

                    #endregion

                    //Load processed fields into result dictionary
                    data[field.InternalName] = value;
                }
                catch (Exception) {}
            }

            #region Additional Info

            if (itemVersion != null)
            {
                data["_VersionNo"] = itemVersion.VersionId;
                data["_VersionCreated"] = itemVersion.Created;
                data["_VersionCreatedBy"] = new SimpleSPPrincipal(itemVersion.CreatedBy.User);
                data["_VersionItemUrl"] = itemVersion.ListItem.FormUrlDisplay() + "&VersionNo=" + itemVersion.VersionId;

                item = itemVersion.ListItem;
            }

            if (itemEventProperties != null && itemEventProperties.ListItem != null)
            {
                item = itemEventProperties.ListItem;
            }

            if (item != null)
            {
                try
                {
                    data["_ItemUniqueId"] = item.UniqueId;
                    if (item.ParentList.ContainsFieldIntName("Attachments"))
                    {
                        List<string> attachments = new List<string>();
                        foreach (string i in item.Attachments)
                        {
                            attachments.Add(item.Attachments.UrlPrefix + i);
                        }
                        data["_Attachments"] = attachments;
                    }
                }
                catch (Exception) {}

                data["_ItemUrl"] = item.FormUrlDisplay();
                data["_HasUniqueRoleAssignments"] = item.HasUniqueRoleAssignments;
                data["_ParentListID"] = item.ParentList.ID;
                data["_ParentWebID"] = item.Web.ID;
                data["__WLI"] = new WebListItemId(item.Web.ID, item.ParentList.ID, item.ID);
            }

            if (itemEventProperties != null && itemEventProperties.ListItem == null)
            {
                data["_ParentListID"] = itemEventProperties.List.ID;
                data["_ParentWebID"] = itemEventProperties.Web.ID;
            }

            #endregion

            data.Sort();
            return data;
        }
    }
}