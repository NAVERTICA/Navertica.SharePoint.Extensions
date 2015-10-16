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
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web;
using System.Xml;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;

namespace Navertica.SharePoint.Extensions
{
    public static class SPListExtensions
    {
        public const string SPDocumentLibraryClassName = "Microsoft.SharePoint.SPDocumentLibrary";
        public const string SPListClassName = "Microsoft.SharePoint.SPList";

        public static SPListItem AddOrUpdateItem(this SPList list, string[] uniqueKeyFields, IDictionary<string, object> itemMetadata)
        {
            SPListItem result;

            if (uniqueKeyFields != null && uniqueKeyFields.Length > 0)
            {
                List<Q> qlist = new List<Q>();

                if (!uniqueKeyFields.All(itemMetadata.ContainsKey))
                {
                    throw new KeyNotFoundException(string.Format("AddOrUpdateItem on list {0} - unique key fields specifies keys {1}, some were not found in item metadata {2}",
                        list.DefaultViewUrl, uniqueKeyFields.JoinStrings(", "), itemMetadata.Select(kvp => kvp.Key + ": " + kvp.Value).JoinStrings(", ")));
                }

                foreach (string key in uniqueKeyFields)
                {
                    // TODO postavit obecnou funkcionalitu, ktera bude umet stavet query pro ruzne typy policek
                    SPField fld = list.GetFieldByInternalName(key);
                    string fldType = fld.TypeAsString;

                    switch (fldType)
                    {
                        case "Text":
                        case "Note":
                            qlist.Add(new Q(QOp.Equal, QVal.Text, key, itemMetadata[key]));
                            break;
                    }
                }

                Q query = new Q(QJoin.And, qlist.ToArray<QComponent>());

                var existing = list.ProcessItems(item => item,
                    query.ToString());

                if (existing.Count > 1)
                {
                    throw new DuplicateNameException(
                        string.Format(
                            "Non unique items {4} found given following combination: unique fields [{0}] with values {{{1}}} in list {2} for query {3} ",
                            uniqueKeyFields.JoinStrings(", "),
                            itemMetadata.Select(kvp => kvp.Key + ": " + kvp.Value).JoinStrings(", "),
                            list.DefaultViewUrl, query, existing.Select(delegate(object o)
                            {
                                if (o is SPListItem) return ( (SPListItem) o ).FormUrlDisplay();
                                if (o != null) return o.ToString();
                                return "null";
                            }).JoinStrings()));
                }

                result = existing.Count == 0 ? list.AddItem() : (SPListItem) existing.First();
            }
            else
            {
                result = list.AddItem();
            }

            foreach (var kvp in itemMetadata)
            {
                if (kvp.Key.ToLowerInvariant() == "contenttype")
                {
                    SPContentType ctp = list.ContentTypes.GetContentType(kvp.Value.ToString());
                    if (ctp == null)
                        throw new Exception(string.Format("List {0} missing content type name {1}", list.DefaultViewUrl, kvp.Value));
                    result["ContentType"] = ctp;
                }
                string fieldType = list.Fields.GetFieldByInternalName(kvp.Key).TypeAsString;

                // TODO this will be done by .Set extension method, which will take care of which type to use etc.

                switch (fieldType)
                {
                    default:
                        result[kvp.Key] = kvp.Value;
                        break;
                }
            }

            result.Update();

            return result;
        }

        /// <summary>
        /// Gets absolute url of the list with no page at the end
        /// e.g. http://portal/Lists/testList/ or http://portal/testLibrary/Forms/
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static string AbsoluteUrl(this SPList list)
        {
            if (list == null) throw new ArgumentNullException("list");

            return list.ParentWeb.Site.Url + list.DefaultViewUrl.Substring(0, list.DefaultViewUrl.LastIndexOf('/') + 1);
        }

        /// <summary>
        /// Adds the ContentType with specified contentTypeId to the list
        /// </summary>
        /// <param name="list"></param>
        /// <param name="contentTypeId"></param>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="FileNotFoundException"></exception>
        /// <exception cref="SPException"></exception>
        /// <exception cref="UnauthorizedAccessException"></exception>
        public static bool AddContentType(this SPList list, SPContentTypeId contentTypeId)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (contentTypeId == null) throw new ArgumentNullException("contentTypeId");

            return AddContentType(list, contentTypeId.ToString());
        }

        /// <summary>
        /// Adds the field to an
        /// </summary>
        /// <param name="list"></param>
        /// <param name="siteField"></param>
        /// <param name="toAllContentTypes"></param>
        public static void AddSiteSPField(this SPList list, SPField siteField, bool toAllContentTypes)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (siteField == null) throw new ArgumentNullException("siteField");

            list.Fields.Add(siteField);
            SPField listField = list.Fields.GetFieldByInternalName(siteField.InternalName);

            if (toAllContentTypes)
            {
                SPFieldLink fieldLink = new SPFieldLink(listField);
                foreach (SPContentType ct in list.ContentTypes.Cast<SPContentType>().Where(ct => !ct.Sealed))
                {
                    try
                    {
                        ct.FieldLinks.Add(fieldLink);
                        ct.Update();
                    }
                    // ReSharper disable once EmptyGeneralCatchClause
                    catch (Exception) {}
                }

                //list.Fields.AddFieldAsXml(siteField.SchemaXml, true, SPAddFieldOptions.AddToAllContentTypes);
            }

            CultureInfo original = Thread.CurrentThread.CurrentUICulture;

            foreach (CultureInfo cul in list.ParentWeb.SupportedUICultures)
            {
                string title = siteField.TitleResource.GetValueForUICulture(cul);
                string desc = siteField.DescriptionResource.GetValueForUICulture(cul);
                Thread.CurrentThread.CurrentUICulture = cul;

                listField.Title = title;
                listField.Description = desc;
            }

            listField.Update(true);

            Thread.CurrentThread.CurrentUICulture = original;
        }

        /// <summary>
        /// Attaches given event receiver
        /// </summary>
        /// <param name="list"></param>
        /// <param name="assemblyName"></param>
        /// <param name="classNameReceiver"></param>
        /// <param name="type"></param>
        /// <param name="syncType"></param>
        /// <param name="sequence"></param>
        public static void AttachEventReceiver(this SPList list, string assemblyName, string classNameReceiver, SPEventReceiverType type, SPEventReceiverSynchronization syncType = SPEventReceiverSynchronization.Default, int sequence = 10000)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (assemblyName == null) throw new ArgumentNullException("assemblyName");
            if (classNameReceiver == null) throw new ArgumentNullException("classNameReceiver");

            // only if it isn't attached already
            if (list.EventReceivers.Cast<SPEventReceiverDefinition>().Any(evRec =>
                ( evRec.Assembly.ToLowerInvariant() == assemblyName.ToLowerInvariant().Trim() ) &&
                ( evRec.Class.ToLowerInvariant() == classNameReceiver.ToLowerInvariant().Trim() ) &&
                ( evRec.Type == type )))
            {
                return;
            }

            SPEventReceiverDefinition def = list.EventReceivers.Add();
            def.Class = classNameReceiver;
            def.Assembly = assemblyName;
            def.Type = type;
            def.Synchronization = syncType;
            def.SequenceNumber = sequence;

            list.ParentWeb.RunWithAllowUnsafeUpdates(def.Update);
        }

        /// <summary>
        /// Adds the ContentType with specified contentTypeId to the list
        /// </summary>
        /// <param name="list"></param>
        /// <param name="contentTypeId">as string</param>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="FileNotFoundException"></exception>
        /// <exception cref="SPException"></exception>
        /// <exception cref="UnauthorizedAccessException"></exception>
        public static bool AddContentType(this SPList list, string contentTypeId)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (contentTypeId == null) throw new ArgumentNullException("contentTypeId");

            SPContentTypeId newContentTypeId;
            try
            {
                newContentTypeId = new SPContentTypeId(contentTypeId.Trim());
            }
            catch (ArgumentException)
            {
                throw new ArgumentException("contentTypeId '" + contentTypeId + "' is not valid");
            }

            SPContentType newContentType = list.ParentWeb.ContentTypes[newContentTypeId] ?? //find on webs CT
                                           list.ParentWeb.Site.RootWeb.ContentTypes[newContentTypeId]; //find on site CT

            if (newContentType == null)
            {
                throw new FileNotFoundException("Can not find content type with contentTypeId '" + contentTypeId + "' trying to add it to list " + list.AbsoluteUrl());
            }

            if (!list.ContentTypes.Contains(newContentTypeId))
            {
                list.ParentWeb.RunWithAllowUnsafeUpdates(() => list.ContentTypes.Add(newContentType) /*Can throw UnauthorizedAccessException, SPException, ...*/);
                return true;
            }

            return false;
        }

        #region ContainsField

        /// <summary>
        /// Checks whether the list contains all the fields with internal names passed in intFieldNames
        /// </summary>
        /// <param name="list"></param>
        /// <param name="guids">guids the list should contain</param>
        /// <returns></returns>
        public static bool ContainsFieldGuid(this SPList list, IEnumerable<Guid> guids)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (guids == null) throw new ArgumentNullException("guids");

            foreach (Guid guid in guids)
            {
                try
                {
                    // ReSharper disable once UnusedVariable
                    var tmp = list.Fields[guid];
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
        /// <param name="list"></param>
        /// <param name="guid"></param>
        /// <returns></returns>
        public static bool ContainsFieldGuid(this SPList list, Guid guid)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (guid == null) throw new ArgumentNullException("guid");

            return ContainsFieldGuid(list, new[] { guid });
        }

        /// <summary>
        /// Checks whether the list contains all the fields with internal names passed in intFieldNames
        /// </summary>
        /// <param name="list"></param>
        /// <param name="intFieldNames">internal names the list should contain</param>
        /// <returns></returns>
        public static bool ContainsFieldIntName(this SPList list, IEnumerable<string> intFieldNames)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (intFieldNames == null) throw new ArgumentNullException("intFieldNames");

            foreach (string fieldname in intFieldNames)
            {
                try
                {
                    list.Fields.GetFieldByInternalName(fieldname);
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
        /// <param name="list"></param>
        /// <param name="intFieldName"></param>
        /// <returns></returns>
        public static bool ContainsFieldIntName(this SPList list, string intFieldName)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (intFieldName == null) throw new ArgumentNullException("intFieldName");

            return ContainsFieldIntName(list, new[] { intFieldName });
        }

        #endregion

        /// <summary>
        /// Gets relative default view url of the list with no page at the end.
        /// e.g. /Lists/testList/ or /testLibrary/Forms/
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static string DefaultViewUrlNoAspx(this SPList list)
        {
            if (list == null) throw new ArgumentNullException("list");

            string result;
            try
            {
                result = list.DefaultViewUrl.Remove(list.DefaultViewUrl.LastIndexOf("/", StringComparison.InvariantCulture) + 1);
            }
            catch (NullReferenceException)
            {
                result = list.Title + " DefaultViewUrlFailed";
            }

            return result;
        }

        #region Urls

        /// <summary>
        /// Gets url of DisplayForm for given itemId
        /// </summary>
        /// <param name="list"></param>
        /// <param name="itemId"></param>
        /// <returns></returns>
        public static string FormUrlDisplay(this SPList list, int itemId)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (itemId < 1) throw new ArgumentException("itemID < 1");

            return list.ParentWeb.Url + "/_layouts/listform.aspx?PageType=4&ListId={" + list.ID + "}&ID=" + itemId;
        }

        /// <summary>
        /// Gets url of EditForm for given itemID
        /// </summary>
        /// <param name="list"></param>
        /// <param name="itemId"></param>
        /// <returns></returns>
        public static string FormUrlEdit(this SPList list, int itemId)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (itemId < 1) throw new ArgumentException("itemID < 1");

            return list.ParentWeb.Url + "/_layouts/listform.aspx?PageType=6&ListId={" + list.ID + "}&ID=" + itemId;
        }

        /// <summary>
        /// Gets url of NewForm
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static string FormUrlNew(this SPList list)
        {
            if (list == null) throw new ArgumentNullException("list");

            return AbsoluteUrl(list) + "NewForm.aspx";
        }

        #endregion

        #region Get Items functions utilizing the Q class for building CAML queries

        /// <summary>
        /// Returns items containing lookupID in lookup named fieldIntName
        /// </summary>
        /// <param name="list"></param>
        /// <param name="fieldIntName"></param>
        /// <param name="lookupId"></param>
        /// <param name="orderBy">null/empty or internal names of fields to order by - if the firest character is >, descendent ordering will be used, ascending by default</param>
        /// <param name="viewFields"> </param>
        /// <param name="andQuery">Q object containing other query parts that will be added as And to the lookup id query</param>
        /// <param name="rowLimit"></param>
        /// <returns>null in case of error</returns>        
        public static SPListItemCollection GetItemsByLookupId(this SPList list, string fieldIntName, int lookupId, IEnumerable<string> orderBy = null, IEnumerable<string> viewFields = null, Q andQuery = null, int rowLimit = -1)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (fieldIntName == null) throw new ArgumentNullException("fieldIntName");
            if (!list.ContainsFieldIntName(fieldIntName)) throw new SPFieldNotFoundException(list, fieldIntName);

            Q query = new Q(QOp.Equal, QVal.LookupId, fieldIntName, lookupId, orderBy, viewFields);
            if (andQuery != null)
            {
                query = new Q(QJoin.And, query, andQuery);
            }

            return GetItemsQuery(list, query, rowLimit);
        }

        /// <summary>
        /// Returns items containing lookupValue in lookup named fieldIntName
        /// </summary>
        /// <param name="list"></param>
        /// <param name="fieldIntName"></param>
        /// <param name="value"></param>
        /// <param name="orderBy">null/empty or internal names of fields to order by - if the firest character is >, descendent ordering will be used, ascending by default</param>
        /// <param name="viewFields"> </param>
        /// <param name="rowLimit"></param>
        /// <returns>null in case of error</returns>
        public static SPListItemCollection GetItemsByLookupValue(this SPList list, string fieldIntName, string value, IEnumerable<string> orderBy = null, IEnumerable<string> viewFields = null, int rowLimit = -1)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (fieldIntName == null) throw new ArgumentNullException("fieldIntName");
            if (value == null) throw new ArgumentNullException("value");
            if (!list.ContainsFieldIntName(fieldIntName)) throw new SPFieldNotFoundException(list, fieldIntName);

            return GetItemsQuery(list, new Q(QOp.Equal, QVal.Text, fieldIntName, value, orderBy, viewFields), rowLimit);
        }

        /// <summary>
        /// Returns items with value in textfield fieldIntName
        /// </summary>
        /// <param name="list"></param>
        /// <param name="fieldIntName"></param>
        /// <param name="value"></param>
        /// <param name="orderBy">null/empty or internal names of fields to order by - if the firest character is >, descendent ordering will be used, ascending by default</param>
        /// <param name="viewFields"> </param>
        /// <param name="rowLimit"></param>
        /// <returns>null in case of error</returns>
        public static SPListItemCollection GetItemsByTextField(this SPList list, string fieldIntName, string value, IEnumerable<string> orderBy = null, IEnumerable<string> viewFields = null, int rowLimit = -1)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (fieldIntName == null) throw new ArgumentNullException("fieldIntName");
            if (value == null) throw new ArgumentNullException("value");
            if (!list.ContainsFieldIntName(fieldIntName)) throw new SPFieldNotFoundException(list, fieldIntName);

            return GetItemsQuery(list, new Q(QOp.Equal, QVal.Text, fieldIntName, value, orderBy, viewFields), rowLimit);
        }

        /// <summary>
        /// Returns items where textfield fieldIntName begins with value
        /// </summary>
        /// <param name="list"></param>
        /// <param name="fieldIntName"></param>
        /// <param name="value"></param>
        /// <param name="orderBy">null/empty or internal names of fields to order by - if the firest character is >, descendent ordering will be used, ascending by default</param>
        /// <param name="viewFields"> </param>
        /// <param name="rowLimit"></param>
        /// <returns>null in case of error</returns>
        public static SPListItemCollection GetItemsByTextFieldBeginsWith(this SPList list, string fieldIntName, string value, IEnumerable<string> orderBy = null, IEnumerable<string> viewFields = null, int rowLimit = -1)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (fieldIntName == null) throw new ArgumentNullException("fieldIntName");
            if (value == null) throw new ArgumentNullException("value");
            if (!list.ContainsFieldIntName(fieldIntName)) throw new SPFieldNotFoundException(list, fieldIntName);

            return GetItemsQuery(list, new Q(QOp.BeginsWith, QVal.Text, fieldIntName, value, orderBy, viewFields), rowLimit);
        }

        /// <summary>
        /// Returns items which contain value in textfield fieldIntName
        /// </summary>
        /// <param name="list"></param>
        /// <param name="fieldIntName"></param>
        /// <param name="value"></param>
        /// <param name="orderBy">null/empty or internal names of fields to order by - if the firest character is >, descendent ordering will be used, ascending by default</param>
        /// <param name="viewFields"> </param>
        /// <param name="rowLimit"></param>
        /// <returns>null in case of error</returns>
        public static SPListItemCollection GetItemsByTextFieldContains(this SPList list, string fieldIntName, string value, IEnumerable<string> orderBy = null, IEnumerable<string> viewFields = null, int rowLimit = -1)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (fieldIntName == null) throw new ArgumentNullException("fieldIntName");
            if (value == null) throw new ArgumentNullException("value");
            if (!list.ContainsFieldIntName(fieldIntName)) throw new SPFieldNotFoundException(list, fieldIntName);

            return GetItemsQuery(list, new Q(QOp.Contains, QVal.Text, fieldIntName, value, orderBy, viewFields), rowLimit);
        }

        /// <summary>
        /// Returns item collection for given Q or null 
        /// </summary>
        /// <param name="list"></param>
        /// <param name="que"></param>
        /// <param name="rowLimit">if > 1, sets the SPQuery RowLimit</param>
        /// <returns></returns>
        public static SPListItemCollection GetItemsQuery(this SPList list, Q que, int rowLimit = -1)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (que == null) throw new ArgumentNullException("que");

            return GetItemsQuery(list, que.ToString(), rowLimit);
        }

        /// <summary>
        /// Returns item collection for given query string or null
        /// </summary>
        /// <param name="list"></param>
        /// <param name="querystr"></param>
        /// <param name="rowLimit">if > 1, sets the SPQuery RowLimit</param>
        /// <returns></returns>
        public static SPListItemCollection GetItemsQuery(this SPList list, string querystr, int rowLimit = -1)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (querystr == null) throw new ArgumentNullException("querystr");

            SPQuery query = GetSPQuery(list, querystr, rowLimit);

            SPListItemCollection coll = list.GetItems(query);

            return coll;
        }

        public static SPQuery GetSPQuery(this SPList list, string querystr, int? rowLimit = -1)
        {
            if (list == null) throw new ArgumentNullException("list");

            SPQuery query = new SPQuery();

            if (!string.IsNullOrEmpty(querystr))
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml("<START>" + querystr + "</START>"); //musi byt obalenm, protoze muze nastat varianta, ze budou dva rootove elementy a ty pak jako XML nenacte a hodi exc
                XmlElement element = doc.DocumentElement;

                #region Check internal names

                List<string> internalFieldNames = new List<string>();

                if (element == null) throw new XmlException("Cannot load CAML query as xml\n\n" + querystr);

                XmlNodeList nodes = element.GetElementsByTagName("FieldRef");
                for (int i = 0; i < nodes.Count; i++)
                {
                    XmlNode node = nodes.Item(i);
                    if (node != null && node.Attributes != null)
                    {
                        string internalName = node.Attributes[0].Value;
                        internalFieldNames.Add(internalName);
                    }
                    else
                    {
                        throw new XmlException("Cannot load node.Attributes for query \n\n" + querystr);
                    }
                }

                // check if all fields in the query exist
                if (!list.ContainsFieldIntName(internalFieldNames)) throw new SPFieldNotFoundException(list, internalFieldNames.Distinct());

                #endregion

                querystr = element.ChildNodes[0].OuterXml.Replace("<Where>", "").Replace("</Where>", "");

                string orderBy = element.ChildNodes.Count > 1 ? element.ChildNodes[1].OuterXml : "";
                string viewFields = element.ChildNodes.Count > 2 ? element.ChildNodes[2].OuterXml : "";
                if (orderBy.StartsWith("<ViewFields>"))
                {
                    viewFields = orderBy;
                    orderBy = "";
                }

                query = new SPQuery { Query = "<Where>" + querystr + "</Where>" + orderBy };

                if (!string.IsNullOrEmpty(viewFields))
                {
                    query.ViewFields = viewFields.Replace("<ViewFields>", "").Replace("</ViewFields>", "");
                    query.ViewFieldsOnly = true;
                }
            }

            if (rowLimit != null && rowLimit > 0)
            {
                query.RowLimit = (uint) rowLimit;
            }

            return query;
        }

        #endregion

        /// <summary>
        /// Returns the field with the specified internal name from the collection.
        /// </summary>
        /// <param name="list"></param>
        /// <param name="strName"></param>
        /// <returns></returns>
        public static SPField GetFieldByInternalName(this SPList list, string strName)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (strName == null) throw new ArgumentNullException("strName");
            if (strName.Trim() == string.Empty) throw new ArgumentException("strName empty");

            return list.Fields.GetFieldByInternalName(strName.Trim());
        }

        /// <summary>
        /// Returns SPFolder for a given relative url
        /// </summary>
        /// <param name="list"></param>
        /// <param name="folderRelativeUrl"></param>
        /// <returns></returns>
        public static SPFolder GetFolder(this SPList list, string folderRelativeUrl)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (folderRelativeUrl == null) throw new ArgumentNullException("folderRelativeUrl");
            if (folderRelativeUrl.Trim() == string.Empty) throw new ArgumentException("folderRelativeUrl empty");

            string listInternalName = list.InternalName();
            string folderPath = folderRelativeUrl;
            string webRelativeUrl = list.ParentWeb.ServerRelativeUrl.Remove(0, 1);

            if (webRelativeUrl != "" && ( folderRelativeUrl.StartsWith(webRelativeUrl) || folderRelativeUrl.StartsWith("/" + webRelativeUrl) ))
            {
                folderPath = folderRelativeUrl.StartsWith("/") ? folderRelativeUrl.Replace("/" + webRelativeUrl, "") : folderRelativeUrl.Replace(webRelativeUrl, "");
            }

            folderPath = folderPath.StartsWith("/") ? folderPath.Remove(0, 1) : folderPath; //odstranit prvni lomitko

            folderPath = folderPath.Replace("Lists/", "");

            if (folderPath.StartsWith(listInternalName)) folderPath = folderPath.Replace(listInternalName, "");

            folderPath = folderPath.StartsWith("/") ? listInternalName + folderPath : listInternalName + "/" + folderPath;
            if (list.GetType().FullName == SPListClassName) folderPath = "Lists/" + folderPath;

            SPFolder folder = list.ParentWeb.GetFolder(folderPath);
            return folder.Exists ? folder : null;
        }

        public static SPFolder GetOrCreateFolder(this SPList list, string path)
        {
            return list.RootFolder.GetOrCreateFolder(path);
        }

        public static bool IsSPDocumentLibrary(this SPList list)
        {
            if (list == null) throw new ArgumentNullException("list");

            return list.GetType() == typeof (SPDocumentLibrary);
        }

        public static bool IsSPList(this SPList list)
        {
            if (list == null) throw new ArgumentNullException("list");

            return list.GetType() == typeof (SPList);
        }

        /// <summary>
        /// Gets a list of internal field names
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static IEnumerable<string> InternalFieldNames(this SPList list)
        {
            if (list == null) throw new ArgumentNullException("list");

            return list.Fields.InternalFieldNames();
        }

        /// <summary>
        /// Gets an "internal name" for the list from its url
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static string InternalName(this SPList list)
        {
            if (list == null) throw new ArgumentNullException("list");

            return ListInternalName(list.DefaultViewUrl);
        }

        /// <summary>
        /// Gets an "internal name" from given list url
        /// </summary>
        /// <param name="fullUrl"></param>
        /// <returns></returns>
        public static string ListInternalName(string fullUrl)
        {
            if (string.IsNullOrWhiteSpace(fullUrl)) return null;

            string url = HttpUtility.UrlDecode(fullUrl) ?? "";
            string listName = string.Empty;

            string[] rets = url.Split('/');

            if (rets.Length == 1) return url;

            if (listName == "" && url.Contains("Forms/Forms")) listName = "Forms"; //muze se stat
            if (listName == "" && url.Contains("Lists/Lists")) listName = "Lists";

            if (!url.ContainsAny(new[] {"Lists", "Forms"}))
            {
                if (rets.Last().Contains(".aspx"))
                {
                    return rets[rets.Length - 2];
                }
                return rets[rets.Length - 1];
            }

            for (int i = 0; i < rets.Length; i++)
            {
                if (rets[i] == "Lists")
                {
                    listName = rets[i + 1];
                    break;
                }
                if (rets[i] == "Forms")
                {
                    listName = rets[i - 1];
                    break;
                }
            }

            return listName;
        }

        /// <summary>
        /// Allow or disable link to List on QuickLaunch Navigation
        /// </summary>
        /// <param name="list"></param>
        /// <param name="state"></param>
        public static void OnOnQuickLaunch(this SPList list, bool state)
        {
            if (list == null) throw new ArgumentNullException("list");

            if (state) //enable
            {
                if (!list.OnQuickLaunch)
                {
                    list.OnQuickLaunch = true;
                    list.Update();
                }
            }
            else //disable
            {
                if (list.OnQuickLaunch)
                {
                    list.OnQuickLaunch = false;
                    list.Update();
                }
            }
        }

        public static SPField OpenField(this SPList list, string fieldIdentification, bool throwExc = true)
        {
            if (fieldIdentification == null) throw new ArgumentNullException("fieldIdentification");

            if (ContainsFieldIntName(list, fieldIdentification))
            {
                return GetFieldByInternalName(list, fieldIdentification);
            }

            Guid fieldGuid = Guid.Empty;
            try
            {
                fieldGuid = new Guid(fieldIdentification);
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch {}

            if (!fieldGuid.IsEmpty())
            {
                try
                {
                    return list.Fields[fieldGuid];
                }
                catch
                {
                    if (throwExc) throw new SPFieldNotFoundException(list, fieldGuid);
                }
            }

            if (throwExc)
            {
                throw new SPFieldNotFoundException(list, fieldIdentification);
            }

            return null;
        }

        #region Process Functions

        /// <summary>
        /// Proccess all items including all folders and subfolders and all items within
        /// </summary>
        /// <param name="list"></param>
        /// <param name="func"></param>
        /// <returns></returns>
        public static ICollection<object> ProcessAllItems(this SPList list, Func<SPListItem, object> func)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (func == null) throw new ArgumentNullException("func");

            return list.RootFolder.ProcessAllItems(func);
        }

        /// <summary>
        /// Executes function on items returned for given querystring, returns a list of values returned by the function
        /// </summary>
        /// <param name="list"></param>
        /// <param name="querystr"></param>
        /// <param name="func"></param>
        /// <param name="rowLimit">Should be set to 1 in case items are to be deleted, otherwise Collection modified exception will occur</param>
        /// <param name="includeSubFolderItems"> </param>
        /// <returns></returns>        
        public static ICollection<object> ProcessItems(this SPList list, Func<SPListItem, object> func, string querystr = "", int rowLimit = -1, bool includeSubFolderItems = true)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (querystr == null) throw new ArgumentNullException("querystr");
            if (func == null) throw new ArgumentNullException("func");

            using (new SPMonitoredScope("ProcessItems"))
            {
                SPQuery q = GetSPQuery(list, querystr, rowLimit);
                if (includeSubFolderItems)
                {
                    q.ViewAttributes = "Scope=\"Recursive\"";
                }

                return ProcessItems(list, func, q);
            }
        }

        public static ICollection<object> ProcessItems(this SPList list, Func<SPListItem, object> func, SPQuery query)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (query == null) throw new ArgumentNullException("query");
            if (func == null) throw new ArgumentNullException("func");

            List<object> result = new List<object>();

            do
            {
                SPListItemCollection coll = list.GetItems(query);

                foreach (SPListItem item in coll)
                {
                    try
                    {
                        result.Add(item.ProcessItem(func));
                    }
                    catch (TerminateException)
                    {
                        return result;
                    }
                    catch (Exception ee)
                    {
                        result.Add(new Exception("error processing item " + item.ID + " " + ( item.Title ?? item.Name ) + "\n" + ee));
                    }
                }

                query.ListItemCollectionPosition = coll.ListItemCollectionPosition;
            }
            while (query.ListItemCollectionPosition != null);

            return result;
        }

        /// <summary>
        /// Executes the function on a given list
        /// </summary>
        /// <param name="list"></param>
        /// <param name="func"></param>
        /// <returns>Result of delegate</returns>
        public static object ProcessList(this SPList list, Func<SPList, object> func)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (func == null) throw new ArgumentNullException("func");

            object result = null;
            list.ParentWeb.RunWithAllowUnsafeUpdates(() => result = func(list)); //at to nemusime porad nastavovat v delegatovi
            return result;
        }

        #endregion

        /// <summary>
        /// Tries to removes the specified event receiver from the list
        /// </summary>
        /// <param name="list"></param>
        /// <param name="assemblyName"></param>
        /// <param name="classNameReceiver"></param>
        /// <param name="type"></param>
        public static void RemoveEventReceiver(this SPList list, string assemblyName, string classNameReceiver, SPEventReceiverType type)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (assemblyName == null) throw new ArgumentNullException("assemblyName");
            if (classNameReceiver == null) throw new ArgumentNullException("classNameReceiver");

            SPEventReceiverDefinition def = list.EventReceivers
                .Cast<SPEventReceiverDefinition>().SingleOrDefault(
                    evRec => ( evRec.Assembly.ToLowerInvariant() == assemblyName.ToLowerInvariant().Trim() ) &&
                             ( evRec.Class.ToLowerInvariant() == classNameReceiver.ToLowerInvariant().Trim() ) &&
                             ( evRec.Type == type ));

            if (def != null)
            {
                list.ParentWeb.RunWithAllowUnsafeUpdates(def.Delete);
            }
        }

        /// <summary>
        /// Tries to change the "internal name" of the list (which is contained in the URL)
        /// </summary>
        /// <param name="list"></param>
        /// <param name="newInternalName"></param>
        /// <returns></returns>
        public static SPList RenameInternalName(this SPList list, string newInternalName)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (newInternalName == null) throw new ArgumentNullException("newInternalName");

            RunElevated(list, delegate(SPList elevList)
            {
                elevList.RootFolder.MoveTo(elevList.RootFolder.Url.Replace(elevList.InternalName(), newInternalName));
                return null;
            });

            return list.ParentWeb.Lists[list.ID];
        }

        /// <summary>
        /// Run func on list as elevated admin
        /// </summary>
        /// <param name="list"></param>
        /// <param name="func"></param>
        /// <returns></returns>
        public static object RunElevated(this SPList list, Func<SPList, object> func)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (func == null) throw new ArgumentNullException("func");

            Guid siteGuid = list.ParentWeb.Site.ID;
            Guid webGuid = list.ParentWeb.ID;
            Guid listGuid = list.ID;
            object result = null;

            if (list.ParentWeb.InSandbox() || list.ParentWeb.CurrentUser.IsSiteAdmin) //Don't runeleveted if user is admin or already run with admin rights
            {
                bool origSiteunsafe = list.ParentWeb.Site.AllowUnsafeUpdates;
                list.ParentWeb.Site.AllowUnsafeUpdates = true;
                list.ParentWeb.RunWithAllowUnsafeUpdates(delegate { result = func(list); });
                list.ParentWeb.Site.AllowUnsafeUpdates = origSiteunsafe;
            }
            else
            {
                using (new SPMonitoredScope("RunElevated SPList - " + list.DefaultViewUrlNoAspx()))
                {
                    SPSecurity.RunWithElevatedPrivileges(delegate
                    {
                        using (SPSite elevatedSite = new SPSite(siteGuid, list.ParentWeb.Site.GetSystemToken()))
                        {
                            elevatedSite.AllowUnsafeUpdates = true;
                            using (SPWeb elevatedWeb = elevatedSite.OpenW(webGuid, true))
                            {
                                elevatedWeb.AllowUnsafeUpdates = true;
                                SPList elevatedList = elevatedWeb.Lists[listGuid];
                                result = func(elevatedList);
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