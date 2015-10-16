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
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using System.Web.Caching;
using System.Web;

namespace Navertica.SharePoint.Extensions
{
    // ReSharper disable once PartialTypeWithSinglePart
    public static partial class SPListItemExtensions
    {
     
        /// <summary>
        /// Copy any item to a target list's root folder.
        /// </summary>
        /// <param name="item">Original item</param>
        /// <param name="targetList">Target list</param>
        /// <param name="deleteOriginal">True to delete original item after successful copy</param>
        /// <param name="overwriteExisting">True to overwrite existing - if no queryStr is used, will look for existing items using filename in libraries and Title in lists</param>
        /// <param name="additional">Optional additional values to set in the copied item - keys are field internal names</param>
        /// <param name="queryStr">Optional CAML query string to find existing item(s), and overwrite the first one</param>
        /// <returns>Copied SPListItem or null</returns>
        public static SPListItem CopyTo(this SPListItem item, SPList targetList, bool deleteOriginal = false, bool overwriteExisting = false, DictionaryNVR additional = null, string queryStr = "")
        {
            if (item == null) throw new ArgumentNullException("item");
            if (targetList == null) throw new ArgumentNullException("targetList");

            return CopyTo(item, targetList.RootFolder, deleteOriginal, overwriteExisting, additional, queryStr);
        }

        /// <summary>
        /// Copy any item to a target list's root folder - complete with metadata and attachments. Fields have to have matching InternalNames and similar enough types.
        /// </summary>
        /// <param name="item">Original item</param>
        /// <param name="toFolder">Target folder</param>
        /// <param name="deleteOriginal">True to delete original item after successful copy</param>
        /// <param name="overwriteExisting">True to overwrite existing - if no queryStr is used, will look for existing item using filename in libraries and Title in lists</param>
        /// <param name="additional">Optional additional metadata fields to set in the copied item - keys are field internal names</param>
        /// <param name="queryStr">Optional CAML query string to find existing item and overwrite it, if there are more then one, ConstraintException will be thrown</param>
        /// <returns></returns>
        public static SPListItem CopyTo(this SPListItem item, SPFolder toFolder, bool deleteOriginal = false, bool overwriteExisting = false, DictionaryNVR additional = null, string queryStr = "")
        {
            if (item == null) throw new ArgumentNullException("item");
            if (toFolder == null) throw new ArgumentNullException("toFolder");

            using (new SPMonitoredScope(item.ID + " - " + item.Title + "CopyTo() - " + toFolder.ServerRelativeUrl))
            {
                SPListItem newItem = null;
                bool itemHasToBeUpdated = false;

                SPList sourceList = item.ParentList;
                SPList targetList = toFolder.ParentWeb.OpenList(toFolder.ParentListId, true);
                string sourceType = sourceList.GetType().FullName;
                string targetType = targetList.GetType().FullName;

                if (sourceType != targetType)
                    throw new ArgumentException("Can not copy item of type item " + sourceType + " to list of type " + targetType);
                if (additional != null && !targetList.ContainsFieldIntName(additional.Keys))
                    throw new SPFieldNotFoundException(targetList, additional.Keys);

                #region create item or find existing to overwrite

                // try to find existing item(s) if there's no query specified
                if (overwriteExisting)
                {
                    SPListItemCollection col;
                    if (string.IsNullOrEmpty(queryStr))
                    {
                        if (item.File != null) // in doc library use filename
                        {
                            col = targetList.GetItemsByTextField("FileLeafRef", item.Name);
                        }
                        else // custom list - use title
                        {
                            col = targetList.GetItemsByTextField("Title", item.Title);
                        }
                    }
                    else
                    {
                        // when there is query, use query
                        col = targetList.GetItemsQuery(queryStr);
                        if (col == null)
                        {
                            throw new FileNotFoundException("by " + queryStr + " was not selected any Item!");
                        }
                    }
                    if (col != null)
                    {
                        if (col.Count == 1) newItem = col[0];
                        else if (col.Count > 1) throw new ConstraintException("Non unique list item found");
                    }
                }

                if (newItem == null || item.File != null)
                {
                    toFolder.ParentWeb.RunWithAllowUnsafeUpdates(delegate
                    {
                        if (item.Folder != null) //folder
                        {
                            itemHasToBeUpdated = true;

                            string folderName = toFolder.ServerRelativeUrl + "/" + item.Folder.Name !=
                                                item.Folder.ServerRelativeUrl
                                ? item.Folder.Name
                                : "Duplicate - " + item.Folder.Name;
                            newItem = targetList.Items.Add(toFolder.ServerRelativeUrl, SPFileSystemObjectType.Folder, folderName);
                        }
                        else if (item.File != null) //doc library file
                        {
                            itemHasToBeUpdated = false;

                            string folderName = item.File.ServerRelativeUrl.Replace(item.File.Name, "");
                            SPFolder fileFolder = item.ParentList.GetFolder(folderName);
                            string filename = toFolder.UniqueId != fileFolder.UniqueId
                                ? item.File.Name
                                : "Duplicate - " + item.File.Name; // copying to the same folder

                            newItem = toFolder.CreateOrUpdateDocument(filename, item.File.OpenBinary(), overwriteExisting);
                        }
                        else // custom list item
                        {
                            itemHasToBeUpdated = true;
                            newItem = targetList.Items.Add(toFolder.ServerRelativeUrl, SPFileSystemObjectType.File, null);
                        }
                    });

                    newItem["ContentTypeId"] = item.ContentType.Parent.Id;

                    // SPListDataValidationException can get thrown here                    
                }

                #endregion

                #region Copy metadata

                // copy metadata - expects fields have the same name and matching (or similar enough) types
                // tries to get values also from hidden and computed fields
                for (int i = 0; i < sourceList.Fields.Count; i++)
                {
                    SPField field;
                    try
                    {
                        field = sourceList.Fields[i];
                    }
                    catch (Exception) // nonvalid field
                    {
                        throw new Exception("Cannot get field from list, probably badly installed nonvalid field");
                    }

                    object fieldValue;

                    try
                    {
                        fieldValue = item[field.InternalName];
                    }
                    catch (Exception) // nonvalid field
                    {
                        throw new SPFieldNotFoundException(sourceList, field.InternalName,
                            "Cannot get value from item. Probably nonvalid field (lookup?): " + field.InternalName);
                    }

                    if (fieldValue == null) continue; // nothing to copy

                    try
                    {
                        if (targetList.ContainsFieldIntName(field.InternalName)
                            && !targetList.Fields.GetFieldByInternalName(field.InternalName).ReadOnlyField
                            && item[field.InternalName] != null
                            && !field.InternalName.EqualAny(new[] { "Attachments", "Order", "FileLeafRef", "MetaInfo" }))
                        {
                            // versioned text field
                            if (field.TypeAsString == "Note" && field.SchemaXml.Contains(@"AppendOnly=""TRUE"""))
                            {
                                if (field.SchemaXml.Contains(@"RichText=""TRUE"""))
                                {
                                    string text = item.GetVersionedMultiLineTextAsHtml(field.InternalName);

                                    if (newItem[field.InternalName] != null &&
                                        newItem[field.InternalName].ToString() != text)
                                    {
                                        itemHasToBeUpdated = true;
                                    }
                                    newItem[field.InternalName] = text;
                                }
                                else
                                {
                                    string text = item.GetVersionedMultiLineTextAsPlainText(field.InternalName);

                                    if (newItem[field.InternalName] != null &&
                                        newItem[field.InternalName].ToString() != text)
                                    {
                                        itemHasToBeUpdated = true;
                                    }
                                    newItem[field.InternalName] = text;
                                }
                            }
                            else
                            {
                                if (newItem[field.InternalName] != null &&
                                    newItem[field.InternalName].GetHashCode() != item[field.InternalName].GetHashCode() &&
                                    field.InternalName != "ContentType")
                                {
                                    itemHasToBeUpdated = true;
                                }
                                // TODO will have to be replaced with Get/Set extensions
                                newItem[field.InternalName] = item[field.InternalName];
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("CopyTo  - Item:" + item.Title + " [" + item.ID + "] field: '" +
                                            field.Title + "' [" + field.InternalName +
                                            "]  problem - perhaps insufficient permissions?\n\n Original Exception:" +
                                            ex);
                    }
                }

                #endregion

                #region Attachment Copy

                if (sourceType == SPListExtensions.SPListClassName)
                {
                    string attName = "";
                    try
                    {
                        foreach (string fileName in item.Attachments)
                        {
                            attName = item.Attachments.UrlPrefix + fileName;

                            SPFile file = sourceList.ParentWeb.GetFile(item.Attachments.UrlPrefix + fileName);

                            bool contains = false;
                            for (int i = 0; i < newItem.Attachments.Count; i++)
                            {
                                if (newItem.Attachments[i] == fileName)
                                {
                                    contains = true;
                                    break;
                                }
                            }
                            if (contains)
                            {
                                SPFile savedFile = sourceList.ParentWeb.GetFile(newItem.Attachments.UrlPrefix + fileName);

                                if (!file.OpenBinary().SequenceEqual(savedFile.OpenBinary()))
                                    // changed attachment 
                                {
                                    newItem.Attachments.Delete(fileName); //delete the existing one

                                    newItem.Attachments.Add(fileName, file.OpenBinary()); //... add the new one
                                    itemHasToBeUpdated = true;
                                }
                            }
                            else
                            {
                                itemHasToBeUpdated = true;
                                newItem.Attachments.Add(fileName, file.OpenBinary());
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Copy attachment problem - " + attName + "\n" + ex);
                    }
                }

                #endregion

                #region Additional Copy

                if (additional != null) // add additional data
                {
                    foreach (KeyValuePair<string, object> kvp in additional)
                    {
                        if (newItem[kvp.Key] == null)
                        {
                            newItem[kvp.Key] = kvp.Value;
                        }
                        else if (newItem[kvp.Key].GetHashCode() != kvp.Value.ToString().GetHashCode())
                        {
                            newItem[kvp.Key] = kvp.Value;
                            itemHasToBeUpdated = true;
                        }
                    }
                }

                #endregion

                if (itemHasToBeUpdated)
                {
                    try
                    {
                        newItem.Web.RunWithAllowUnsafeUpdates(newItem.Update);
                    }
                    catch (Exception exc)
                    {
                        throw new Exception("Update Fail\n\n" + exc);
                    }
                }

                try
                {
                    if (deleteOriginal) item.Web.RunWithAllowUnsafeUpdates(() => item.Recycle()); // delete the original item
                }
                catch (Exception exc)
                {
                    throw new Exception("CopyTo - problem deleting the original item - user " +
                                        item.ParentList.ParentWeb.CurrentUser.Name +
                                        " might not have permissions to delete in list " + item.ParentList.DefaultViewUrl + "\n" +
                                        exc);
                }

                if (item.IsFolder())
                {
                    item.Folder.ProcessItems(delegate(SPListItem it)
                    {
                        it.CopyTo(newItem.Folder, deleteOriginal, overwriteExisting, additional, queryStr);
                        return null;
                    });
                }


                return newItem;
            }
        }

        /// <summary>
        /// Delete item under elevated rights
        /// </summary>
        /// <param name="item"></param>
        public static void DeleteItemElevated(this SPListItem item)
        {
            if (item == null) throw new ArgumentNullException("item");

            item.RunElevated(delegate(SPListItem elevItem)
            {
                elevItem.Delete();
                return null;
            });
        }

        /// <summary>
        /// Clean up versions, only keep the last X
        /// </summary>
        /// <param name="item"></param>
        /// <param name="numberOfLastVersions">Number of versions to keep</param>
        public static void DeleteVersionsExceptLast(this SPListItem item, int numberOfLastVersions)
        {
            List<int> ids = item.Versions.Cast<SPListItemVersion>().Select(v => v.VersionId).OrderByDescending(v => v).ToList();

            if (ids.Count > numberOfLastVersions)
            {
                ids.RemoveRange(0, numberOfLastVersions);

                foreach (int id in ids)
                {
                    SPListItemVersion ver = item.Versions.GetVersionFromID(id);
                    ver.Delete();
                }
            }
        }

        /// <summary>
        /// Absolute url for the item's document
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static string DocumentUrl(this SPListItem item)
        {
            if (item == null) throw new ArgumentNullException("item");

            try
            {
                return (string) item[SPBuiltInFieldId.EncodedAbsUrl];
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        ///  List of field internal names that changed between two item versions
        /// </summary>
        /// <param name="newVersion"></param>
        /// <param name="oldVersion"></param>
        /// <returns></returns>
        public static List<string> FieldsChangedBetweenVersions(this SPListItemVersion newVersion, SPListItemVersion oldVersion)
        {
            if (newVersion == null) throw new ArgumentNullException("newVersion");
            if (oldVersion == null) throw new ArgumentNullException("oldVersion");

            using (new SPMonitoredScope("FieldsChangedBetweenVersions"))
            {
                List<string> changedFields = new List<string>();

                for (int i = 0; i < newVersion.ListItem.ParentList.Fields.Count; i++)
                {
                    SPField field;
                    try
                    {
                        field = newVersion.ListItem.ParentList.Fields[i];
                    }
                    catch
                    {
                        continue;
                    }

                    if (newVersion[field.InternalName] != oldVersion[field.InternalName])
                    {
                        changedFields.Add(field.InternalName);
                    }
                }

                return changedFields;
            }
        }

        /// <summary>
        /// Get absolute url of item's DisplayForm 
        /// </summary>
        /// <param name="item"></param>
        /// <param name="useRelativeUrl"> </param>
        /// <returns></returns>
        public static string FormUrlDisplay(this SPListItem item, bool useRelativeUrl = false)
        {
            if (item == null) throw new ArgumentNullException("item");

            return ( useRelativeUrl ? "" + "/" : item.Web.Url + "/" ) + item.ParentList.Forms[PAGETYPE.PAGE_DISPLAYFORM].Url + "?ID=" + item.ID;
        }

        /// <summary>
        /// Get absolute url of item's EditForm 
        /// </summary>
        /// <param name="item"></param>
        /// <param name="useRelativeUrl"> </param>
        /// <returns></returns>
        public static string FormUrlEdit(this SPListItem item, bool useRelativeUrl = false)
        {
            if (item == null) throw new ArgumentNullException("item");

            return ( useRelativeUrl ? "" + "/" : item.Web.Url + "/" ) + item.ParentList.Forms[PAGETYPE.PAGE_EDITFORM].Url + "?ID=" + item.ID;
        }

        /// <summary>
        /// Gets SPFolder with attachments for item
        /// </summary>
        /// <param name="item"></param>
        /// <returns>SPFolder</returns>
        public static SPFolder GetAttachments(this SPListItem item)
        {
            if (item == null) throw new ArgumentNullException("item");

            return item.ParentList.ParentWeb
                .Folders["Lists"]
                .SubFolders[item.ParentList.DefaultViewUrl.Substring(0, item.ParentList.DefaultViewUrl.LastIndexOf("/", StringComparison.InvariantCulture))]
                .SubFolders["Attachments"]
                .SubFolders[item.ID.ToString(item.Web.UICulture)];
        }

        public static string GetDateTimeString(this SPListItem item, string internalName, SPUser user, bool incudeTime = false)
        {
            DateTime? date = (DateTime?) item[internalName];
            if (date == null) return "";

            return ( (DateTime) date ).ToStringLocalized(incudeTime, user.GetPreferredLanguage());
        }

        #region Bound lookups


        /// <summary>
        /// Returns collection of items from listWithLookups, which point to the item in their boundLookupField
        /// </summary>
        /// <param name="item"></param>
        /// <param name="boundLookupField"></param>
        /// <param name="orderBy">null/empty or internal names of fields to order by - if the firest character is >, descendent ordering will be used, ascending by default</param>
        /// <returns></returns>
        public static SPListItemCollection GetBoundLookupItems(this SPListItem item, SPFieldLookup boundLookupField, IEnumerable<string> orderBy = null)
        {
            if (item == null) throw new ArgumentNullException("item");
            if (boundLookupField == null) throw new ArgumentNullException("boundLookupField");

            if (!boundLookupField.ParentList.ContainsFieldIntName(boundLookupField.InternalName)) return null;

            SPListItemCollection results = boundLookupField.ParentList.GetItemsByLookupId(boundLookupField.InternalName, item.ID, orderBy);

            return results;
        }

        #endregion

        #region Lookup Functions

        /// <summary>
        /// Returns identification of a lookup's target list 
        /// </summary>
        /// <param name="item"></param>
        /// <param name="lookupIntName"></param>
        /// <returns>WebListId</returns>
        /// <exception cref="SPException"></exception>
        /// <exception cref="SPFieldNotFoundException"></exception>
        public static WebListId GetLookupList(this SPListItem item, string lookupIntName)
        {
            if (item == null) throw new ArgumentNullException("item");
            if (lookupIntName == null) throw new ArgumentNullException("lookupIntName");

            SPList list = item.ParentList;

            if (!list.ContainsFieldIntName(lookupIntName)) throw new SPFieldNotFoundException(list, lookupIntName);

            SPField lookupField = list.GetFieldByInternalName(lookupIntName);
            if (!lookupField.IsLookup()) throw new SPException("Field " + lookupIntName + " is not Lookup");

            WebListId result = new WebListId();

            if (lookupIntName.EqualAny(new[] { "FileRef", "ItemChildCount", "FolderChildCount" }))
            {
                result.InvalidMessage = "Get lookup form " + lookupIntName + "is not allowed";
                return result;
            }

            try
            {
                using (SPWeb lookupWeb = list.ParentWeb.Site.OpenW(( (SPFieldLookup) lookupField ).LookupWebId, true))
                {
                    try
                    {
                        SPList lookupList = lookupWeb.OpenList(((SPFieldLookup) lookupField).LookupList, true);
                        result = new WebListId(lookupList);
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
        /// Gets identification of the item a lookup is pointing at 
        /// </summary>
        /// <param name="item"></param>
        /// <param name="lookupIntName"></param>
        /// <returns>WebListItemId</returns>
        /// <exception cref="SPListItemNotFoundException"></exception>
        /// <exception cref="SPFieldNotFoundException"></exception>
        public static WebListItemId GetItemFromLookup(this SPListItem item, string lookupIntName)
        {
            if (item == null) throw new ArgumentNullException("item");
            if (lookupIntName == null) throw new ArgumentNullException("lookupIntName");

            using (new SPMonitoredScope("GetItemFromLookup"))
            {
                WebListItemId result = new WebListItemId();
                WebListId webListId = GetLookupList(item, lookupIntName);

                if (webListId.IsValid)
                {
                    using (SPWeb lookupWeb = item.Web.Site.OpenW(webListId.WebGuid, true))
                    {
                        SPList lookupList = lookupWeb.Lists[webListId.ListGuid];

                        int id = ( item[lookupIntName] ?? String.Empty ).ToString().GetLookupIndex();

                        if (id > 0)
                        {
                            try
                            {
                                result = new WebListItemId(lookupList.GetItemById(id));
                            }
                            catch
                            {
                                SPListItemNotFoundException exc = new SPListItemNotFoundException(id, lookupList);
                                result.InvalidMessage = exc.Message;
                                return result;
                            }
                        }
                        else
                        {
                            result.InvalidMessage = "id < 1 - lookup is null or empty";
                            return result;
                        }
                    }
                }
                else
                {
                    result.InvalidMessage = webListId.InvalidMessage;
                }

                return result;
            }
        }

        /// <summary>
        /// Returns WebListItemIds with info of items contained in originalItem's lookup 
        /// </summary>
        /// <param name="item"></param>
        /// <param name="lookupIntName"></param>
        /// <returns>List of WebListItemId</returns>
        public static WebListItemDictionary GetItemsFromLookup(this SPListItem item, string lookupIntName)
        {
            if (item == null) throw new ArgumentNullException("item");
            if (lookupIntName == null) throw new ArgumentNullException("lookupIntName");

            using (new SPMonitoredScope("GetItemsFromLookup"))
            {
                WebListItemDictionary result = new WebListItemDictionary();
                WebListId webListId = GetLookupList(item, lookupIntName);

                if (webListId.IsValid)
                {
                    using (SPWeb lookupWeb = item.Web.Site.OpenW(webListId.WebGuid, true))
                    {
                        SPList lookupList = lookupWeb.Lists[webListId.ListGuid];

                        int[] ids = ( item[lookupIntName] ?? "" ).ToString().GetLookupIndexes();

                        foreach (int itemId in ids.OrderBy(i => i))
                        {
                            if (itemId > 0)
                            {
                                try
                                {
                                    result.Add(new WebListItemId(lookupList.GetItemById(itemId)));
                                }
                                // ReSharper disable once EmptyGeneralCatchClause
                                catch {}
                            }
                        }
                    }
                }

                return result;
            }
        }

        /// <summary>
        /// Returns WebListItemIds with info of items contained in originalItem version's lookup 
        /// </summary>
        /// <param name="itemVersion"></param>
        /// <param name="lookupIntName"></param>
        /// <returns>List of WebListItemId</returns>
        public static WebListItemDictionary GetItemsFromLookup(this SPListItemVersion itemVersion, string lookupIntName)
        {
            if (itemVersion == null) throw new ArgumentNullException("itemVersion");
            if (lookupIntName == null) throw new ArgumentNullException("lookupIntName");

            using (new SPMonitoredScope("GetItemsFromLookup"))
            {
                WebListItemDictionary result = new WebListItemDictionary();
                WebListId webListId = GetLookupList(itemVersion.ListItem, lookupIntName);

                if (webListId.IsValid)
                {
                    using (SPWeb lookupWeb = itemVersion.ListItem.Web.Site.OpenW(webListId.WebGuid, true))
                    {
                        SPList lookupList = lookupWeb.Lists[webListId.ListGuid];

                        int[] ids = ( itemVersion[lookupIntName] ?? "" ).ToString().GetLookupIndexes();

                        foreach (int itemId in ids.OrderBy(i => i))
                        {
                            if (itemId > 0)
                            {
                                try
                                {
                                    result.Add(new WebListItemId(lookupList.GetItemById(itemId)));
                                }
                                // ReSharper disable once EmptyGeneralCatchClause
                                catch {}
                            }
                        }
                    }
                }

                return result;
            }
        }

        /// <summary>
        /// Executes the function on the list to which a lookup is pointed
        /// </summary>
        /// <param name="item"></param>
        /// <param name="lookupIntName"></param>
        /// <param name="func"></param>
        /// <returns>Result of delegate</returns>
        public static object ProcessLookupList(this SPListItem item, string lookupIntName, Func<SPList, object> func)
        {
            if (item == null) throw new ArgumentNullException("item");
            if (lookupIntName == null) throw new ArgumentNullException("lookupIntName");
            if (func == null) throw new ArgumentNullException("func");
            if (!item.ParentList.ContainsFieldIntName(lookupIntName)) throw new SPFieldNotFoundException(item.ParentList, lookupIntName);

            using (new SPMonitoredScope("ProcessLookupList"))
            {
                object result;

                WebListId listId = GetLookupList(item, lookupIntName);
                if (listId.IsValid)
                {
                    result = listId.ProcessList(item.Web.Site, delegate(SPList lookupList)
                    {
                        return func(lookupList);
                    });
                }
                else
                {
                    throw new SPException(listId.InvalidMessage);
                }

                return result;
            }
        }

        /// <summary>
        /// Executes the function on items the lookup is pointing at, returning the results
        /// </summary>
        /// <param name="item"></param>
        /// <param name="lookupIntName"></param>
        /// <param name="func"></param>
        /// <returns></returns>
        public static ICollection<object> ProcessLookupItems(this SPListItem item, string lookupIntName, Func<SPListItem, object> func)
        {
            if (item == null) throw new ArgumentNullException("item");
            if (lookupIntName == null) throw new ArgumentNullException("lookupIntName");
            if (func == null) throw new ArgumentNullException("func");
            if (!item.ParentList.ContainsFieldIntName(lookupIntName)) throw new SPFieldNotFoundException(item.ParentList, lookupIntName);

            using (new SPMonitoredScope("ProcessLookupItems"))
            {
                List<object> results = new List<object>();
                WebListItemDictionary dict = item.GetItemsFromLookup(lookupIntName);
                dict.ProcessItems(item.Web.Site, delegate(SPListItem lookItem)
                {
                    try
                    {
                        results.Add(func(lookItem));
                    }
                    catch (Exception excFunc)
                    {
                        results.Add(excFunc);
                    }

                    return null;
                });

                return results;
            }
        }

        /// <summary>
        /// Executes the function on HISTORIC VERSIONS of the items the lookup is pointing at, returning the results
        /// </summary>
        /// <param name="itemVersion"></param>
        /// <param name="lookupIntName"></param>
        /// <param name="func"></param>
        /// <returns></returns>
        public static ICollection<object> ProcessLookupItems(this SPListItemVersion itemVersion, string lookupIntName, Func<SPListItemVersion, object> func)
        {
            if (itemVersion == null) throw new ArgumentNullException("itemVersion");
            if (lookupIntName == null) throw new ArgumentNullException("lookupIntName");
            if (func == null) throw new ArgumentNullException("func");
            if (!itemVersion.ListItem.ParentList.ContainsFieldIntName(lookupIntName)) throw new SPFieldNotFoundException(itemVersion.ListItem.ParentList, lookupIntName);

            DateTime versionDateTime = itemVersion.Created.ToUniversalTime();

            using (new SPMonitoredScope("ProcessLookupItems"))
            {
                List<object> results = new List<object>();
                WebListItemDictionary dict = itemVersion.GetItemsFromLookup(lookupIntName);

                if (itemVersion.IsCurrentVersion)
                {
                    dict.ProcessItems(itemVersion.ListItem.Web.Site,
                        delegate(SPListItem lookItem)
                        {
                            try
                            {
                                results.Add(func(lookItem.Versions[0]));
                            }
                            catch (Exception excFunc)
                            {
                                results.Add(excFunc);
                            }

                            return null;
                        });
                }
                else
                {
                    dict.ProcessItemVersion(itemVersion.ListItem.Web.Site, versionDateTime,
                        delegate(SPListItemVersion lookItem)
                        {
                            try
                            {
                                results.Add(func(lookItem));
                            }
                            catch (Exception excFunc)
                            {
                                results.Add(excFunc);
                            }

                            return null;
                        });
                }

                return results;
            }
        }

        #endregion

        public static SPFolder GetParentFolder(this SPListItem item)
        {
            if (item == null) throw new ArgumentNullException("item");

            SPFile spFile = item.Web.GetFile(item.Url);

            return spFile == null ? null : spFile.ParentFolder;
        }

        public static string GetTitleLink(this SPListItem item)
        {
            if (item == null) throw new ArgumentNullException("item");

            return "<a href='" + item.FormUrlDisplay() + "'>" + ( item["Title"] ?? item.Name ) + "</a>";
        }

        /// <summary>
        /// Gets the contents of a versioned text field as HTML
        /// </summary>
        /// <param name="item"></param>
        /// <param name="fieldInternalName"></param>
        /// <returns></returns>
        public static string GetVersionedMultiLineTextAsHtml(this SPListItem item, string fieldInternalName)
        {
            if (item == null) throw new ArgumentNullException("item");
            if (fieldInternalName == null) throw new ArgumentNullException("fieldInternalName");

            return item.Versions[0].GetVersionedMultiLineTextAsHtml(fieldInternalName);
        }

        public static string GetVersionedMultiLineTextAsHtml(this SPListItemVersion itemVersion, string fieldInternalName)
        {
            if (itemVersion == null) throw new ArgumentNullException("itemVersion");
            if (fieldInternalName == null) throw new ArgumentNullException("fieldInternalName");
            var versionsDict = itemVersion.GetVersionedMultiLineTextAsSortedDictionary(fieldInternalName);

            StringBuilder sb = new StringBuilder();
            // tahle podivnost se splhanim nahoru a dolu po objektech tu pry je kvuli tomu, ze s vysledkem z SPQuery by to nejelo
            SPFieldMultiLineText field = itemVersion.ListItem.Fields.GetFieldByInternalName(fieldInternalName) as SPFieldMultiLineText;
            if (field == null) throw new SPFieldNotFoundException(itemVersion.ListItem.ParentList, fieldInternalName);

            foreach (KeyValuePair<DateTime, Dictionary<string, string>> kvp in versionsDict)
            {
                if (kvp.Key > itemVersion.Created) continue;
                string comment = kvp.Value["ContentsHTML"];
                if (comment != null && comment.Trim() != string.Empty && comment.Trim() != "<div></div>")
                {
                    sb.Append(kvp.Value["CreatedByName"]).Append(" (");
                    sb.Append(kvp.Key.ToString(itemVersion.ListItem.Web.UICulture));
                    sb.Append("): ");
                    sb.Append(comment);
                    sb.Append("\n\r");
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// Gets the contents of a versioned text field as text
        /// </summary>
        /// <param name="item"></param>
        /// <param name="fieldInternalName"></param>
        /// <returns></returns>
        public static string GetVersionedMultiLineTextAsPlainText(this SPListItem item, string fieldInternalName)
        {
            if (item == null) throw new ArgumentNullException("item");
            if (fieldInternalName == null) throw new ArgumentNullException("fieldInternalName");

            return item.Versions[0].GetVersionedMultiLineTextAsPlainText(fieldInternalName);
        }

        public static string GetVersionedMultiLineTextAsPlainText(this SPListItemVersion itemVersion, string fieldInternalName)
        {
            if (itemVersion == null) throw new ArgumentNullException("itemVersion");
            if (fieldInternalName == null) throw new ArgumentNullException("fieldInternalName");
            var versionsDict = itemVersion.GetVersionedMultiLineTextAsSortedDictionary(fieldInternalName);

            StringBuilder sb = new StringBuilder();
            // tahle podivnost se splhanim nahoru a dolu po objektech tu pry je kvuli tomu, ze s vysledkem z SPQuery by to nejelo
            SPFieldMultiLineText field = itemVersion.ListItem.Fields.GetFieldByInternalName(fieldInternalName) as SPFieldMultiLineText;
            if (field == null) throw new SPFieldNotFoundException(itemVersion.ListItem.ParentList, fieldInternalName);

            foreach (KeyValuePair<DateTime, Dictionary<string, string>> kvp in versionsDict)
            {
                if (kvp.Key > itemVersion.Created) continue;
                string comment = kvp.Value["Contents"];
                if (comment != null && comment.Trim() != string.Empty)
                {
                    sb.Append(kvp.Value["CreatedByName"]).Append(" (");
                    sb.Append(kvp.Key.ToString(itemVersion.ListItem.Web.UICulture));
                    sb.Append("): ");
                    sb.Append(comment);
                    sb.Append("\n\r");
                }
            }

            return sb.ToString();
        }

        public static SortedDictionary<DateTime, Dictionary<string, string>> GetVersionedMultiLineTextAsSortedDictionary(this SPListItem item, string fieldInternalName)
        {
            return item.Versions[0].GetVersionedMultiLineTextAsSortedDictionary(fieldInternalName);
        }

        public static SortedDictionary<DateTime, Dictionary<string, string>> GetVersionedMultiLineTextAsSortedDictionary(this SPListItemVersion itemversion, string fieldInternalName)
        {
            if (itemversion == null) throw new ArgumentNullException("itemversion");
            if (fieldInternalName == null) throw new ArgumentNullException("fieldInternalName");

            string cacheKey = "NVR_MultiLine" + itemversion.ListItem.UniqueId + "_v" + itemversion.VersionId;

            var cached = HttpContext.Current.Cache.Get(cacheKey);
            if (cached != null)
            {
                return new SortedDictionary<DateTime, Dictionary<string, string>>(( (SortedDictionary<DateTime, Dictionary<string, string>>) cached ));
            }

            var resultdict = new SortedDictionary<DateTime, Dictionary<string, string>>();

            // tahle podivnost se splhanim nahoru a dolu po objektech tu pry je kvuli tomu, ze s vysledkem z SPQuery by to nejelo
            var field = itemversion.ListItem.Fields.GetFieldByInternalName(fieldInternalName) as SPFieldMultiLineText;
            if (field == null) throw new SPFieldNotFoundException(itemversion.ListItem.ParentList, fieldInternalName);

            foreach (
                SPListItemVersion version in
                    itemversion.ListItem.Web.Lists[itemversion.ListItem.ParentList.ID]
                        .GetItemByUniqueId(itemversion.ListItem.UniqueId).Versions)
            {
                if (version.Created > itemversion.Created) continue;

                string comment = field.GetFieldValueAsText(version.ListItem[fieldInternalName]);
                if (comment != null && comment.Trim() != string.Empty)
                {
                    string commentHTML = field.GetFieldValueAsHtml(version[fieldInternalName]);
                    resultdict[version.Created] = new Dictionary<string, string>
                    {
                        { "CreatedBy", version.CreatedBy.User.LoginNameNormalized() },
                        { "CreatedByName", version.CreatedBy.User.Name },
                        { "Contents", comment },
                        { "ContentsHTML", commentHTML },
                    };
                }
            }

            HttpContext.Current.Cache.Insert(cacheKey, resultdict, null, DateTime.Now.AddMinutes(2), Cache.NoSlidingExpiration);

            return resultdict;
        }

        /// <summary>
        /// Returns the internal names of all fields in the SPListItem
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static IEnumerable<string> InternalFieldNames(this SPListItem item)
        {
            if (item == null) throw new ArgumentNullException("item");

            return item.Fields.InternalFieldNames();
        }

        /// <summary>
        /// Returns true if the SPListItem is a folder
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static bool IsFolder(this SPListItem item)
        {
            if (item == null) throw new ArgumentNullException("item");

            return item.Folder != null;
        }

        /// <summary>
        /// Checks if given collection is valid. If we load collection by CAML, it might not be null, but there might be no items
        /// </summary>
        /// <param name="col"></param>
        /// <returns></returns>
        public static bool IsValid(this SPListItemCollection col)
        {
            if (col == null) throw new ArgumentNullException("col");

            try
            {
                // test if collection works
                // ReSharper disable UnusedVariable
                int count = col.Count;
                return true;
                // ReSharper restore UnusedVariable
            }
            catch
            {
                return false;
            }
        }

        #region ItemUpdates Functions

        public static void Update(this SPListItem item, bool runReceiver = false)
        {
            if (item == null) throw new ArgumentNullException("item");

            if (runReceiver)
            {
                item.Update();
            }
            else
            {
                item.Web.RunWithAllowUnsafeUpdates(delegate
                {
                    SPEventManagerWrapper.DisableEventFiring();
                    item.Update();
                    SPEventManagerWrapper.EnableEventFiring();
                });
            }
        }

        public static void SystemUpdate(this SPListItem item, bool incrementListItemVersion = false, bool runReceiver = false)
        {
            if (item == null) throw new ArgumentNullException("item");

            if (runReceiver)
            {
                item.SystemUpdate(incrementListItemVersion);
            }
            else
            {
                item.Web.RunWithAllowUnsafeUpdates(delegate
                {
                    SPEventManagerWrapper.DisableEventFiring();
                    item.SystemUpdate(incrementListItemVersion);
                    SPEventManagerWrapper.EnableEventFiring();
                });
            }
        }

        /// <summary>
        /// Like SPListItem.Update(), but shouldn't trigger event receivers. Values to be updated can be passed in a dictionary
        /// where key is the field internal name.
        /// </summary>
        /// <param name="item"></param>
        /// <param name="increaseVersion"></param>
        /// <param name="values"></param>
        public static void ItemSystemUpdate(this SPListItem item, bool increaseVersion = false, DictionaryNVR values = null)
        {
            if (item == null) throw new ArgumentNullException("item");
            if (values != null && !item.ParentList.ContainsFieldIntName(values.Keys)) throw new SPFieldNotFoundException(item.ParentList, values.Keys);

            item.Web.RunWithAllowUnsafeUpdates(delegate
            {
                SPEventManagerWrapper.DisableEventFiring();

                if (values != null)
                {
                    foreach (KeyValuePair<string, object> kvp in values)
                    {
                        item[kvp.Key] = kvp.Value;
                    }
                }

                item.SystemUpdate(increaseVersion);

                SPEventManagerWrapper.EnableEventFiring();
            });
        }

        #endregion

        /// <summary>
        /// Executes the delegate on the item. Returns whatever the delegate returns.
        /// </summary>
        /// <param name="item"></param>
        /// <param name="func"></param>
        /// <returns>result of delegate</returns>
        public static object ProcessItem(this SPListItem item, Func<SPListItem, object> func)
        {
            if (item == null) throw new ArgumentNullException("item");
            if (func == null) throw new ArgumentNullException("func");

            object result = null;
            item.Web.RunWithAllowUnsafeUpdates(() => result = func(item));
            return result;
        }

        /// <summary>
        /// Executes the delegate on the item. Returns whatever the delegate returns.
        /// </summary>
        /// <param name="itemVersion"></param>
        /// <param name="func"></param>
        /// <returns>result of delegate</returns>
        public static object ProcessItem(this SPListItemVersion itemVersion, Func<SPListItemVersion, object> func)
        {
            if (itemVersion == null) throw new ArgumentNullException("itemVersion");
            if (func == null) throw new ArgumentNullException("func");

            object result = null;
            itemVersion.ListItem.Web.RunWithAllowUnsafeUpdates(() => result = func(itemVersion));
            return result;
        }

        public static ICollection<object> ProcessItems(this SPListItemCollection collection, Func<SPListItem, object> func)
        {
            if (collection == null) throw new ArgumentNullException("collection");
            if (func == null) throw new ArgumentNullException("func");

            using (new SPMonitoredScope("ProcessItems : " + collection.Count))
            {
                List<object> result = collection.Cast<SPListItem>().Select(i => i.ProcessItem(func)).ToList();
                return result;
            }
        }

        /// <summary>
        /// Recycle item under elevated rights
        /// </summary>
        /// <param name="item"></param>
        public static void RecycleItemElevated(this SPListItem item)
        {
            if (item == null) throw new ArgumentNullException("item");

            item.RunElevated(delegate(SPListItem elevItem)
            {
                elevItem.Recycle();
                return null;
            });
        }

        /// <summary>
        /// Reopens the item under elevated privilages and executes the delegate on the elevated item. Returns whatever the delegate returns.
        /// </summary>
        /// <param name="item"></param>
        /// <param name="func">ExecuteOnListItem(SPListItem) delegate</param>
        /// <returns>result of delegate</returns>
        public static object RunElevated(this SPListItem item, Func<SPListItem, object> func)
        {
            if (item == null) throw new ArgumentNullException("item");
            if (func == null) throw new ArgumentNullException("func");

            Guid siteGuid = item.Web.Site.ID;
            Guid webGuid = item.Web.ID;
            Guid listGuid = item.ParentList.ID;
            int itemId = item.ID;
            object result = null;

            if (item.Web.InSandbox() || item.Web.CurrentUser.IsSiteAdmin) //Don't run eleveted if user is admin or already run with admin rights
            {
                bool origSiteUnsafe = item.Web.Site.AllowUnsafeUpdates;
                item.Web.Site.AllowUnsafeUpdates = true;
                item.Web.RunWithAllowUnsafeUpdates(delegate { result = func(item); });
                item.Web.Site.AllowUnsafeUpdates = origSiteUnsafe;
            }
            else
            {
                using (new SPMonitoredScope("RunElevated SPListItem - " + item.ID))
                {
                    SPSecurity.RunWithElevatedPrivileges(delegate
                    {
                        using (SPSite elevatedSite = new SPSite(siteGuid, item.Web.Site.GetSystemToken()))
                        {
                            elevatedSite.AllowUnsafeUpdates = true;
                            using (SPWeb elevatedWeb = elevatedSite.OpenW(webGuid, true))
                            {
                                elevatedWeb.AllowUnsafeUpdates = true;
                                SPListItem elevatedItem = elevatedWeb.Lists[listGuid].GetItemById(itemId);
                                result = func(elevatedItem);
                                elevatedWeb.AllowUnsafeUpdates = false;
                            }
                            elevatedSite.AllowUnsafeUpdates = false;
                        }
                    });
                }
            }

            return result;
        }
    }
}