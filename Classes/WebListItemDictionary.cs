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
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Navertica.SharePoint.Extensions;

// ReSharper disable once CheckNamespace
namespace Navertica.SharePoint
{
    /// <summary>
    ///     For passing and handling SPRequest-independent references to SPListItems grouped together by webs and lists
    /// </summary>
    public class WebListItemDictionary : Dictionary<Guid, Dictionary<Guid, List<int>>>
    {
        public int ListGuidCount
        {
            get { return Keys.SelectMany(g => this[g]).Count(); }
        }

        public int ListItemCount
        {
            get
            {
                return
                    (from webguid in Keys
                        from listGuid in this[webguid].Keys
                        select this[webguid][listGuid]
                        into o
                        select o.Count).Sum();
            }
        }

        public bool Contains(WebListItemId wli)
        {
            foreach (Guid webGuid in Keys)
            {
                if (webGuid == wli.WebGuid)
                {
                    foreach (Guid listGuid in this[webGuid].Keys)
                    {
                        if (listGuid == wli.ListGuid)
                        {
                            foreach (int id in this[webGuid][listGuid])
                            {
                                if (id == wli.Item) return true;
                            }
                        }
                    }
                }
            }
            return false;
        }

        #region Constructors

        public WebListItemDictionary()
        {
        }

        public WebListItemDictionary(WebListItemId id)
        {
            if (id == null) throw new ArgumentNullException();

            Add(id);
        }

        public WebListItemDictionary(IEnumerable<WebListItemId> ids)
        {
            if (ids == null) throw new ArgumentNullException();

            foreach (WebListItemId id in ids)
            {
                Add(id);
            }
        }

        public WebListItemDictionary(WebListItemDictionary webListItemDictionary)
        {
            if (webListItemDictionary == null) throw new ArgumentNullException();

            AddRange(webListItemDictionary);
        }

        /// <summary>
        ///     Z typického SiteConfig stringu ve tvaru "webGuid:listGuid#itemID;itemID, webGuid:listGuid#itemID;itemID" vrátí
        ///     slovník, kde klíčem je guid webu, a hodnotou
        ///     další slovník s guidy seznamů jako klíči a listem integerů jako hodnotou
        /// </summary>
        /// <param name="weblistitemguids">webguid:listguid#1;2;3;4</param>
        /// <returns>
        ///     slovnik, kde klic je Guid webu, a hodnota je slovnik s Guidem seznamu jako klicem a Listem integeru jako ID
        ///     polozek
        /// </returns>
        public WebListItemDictionary(string weblistitemguids)
        {
            if (weblistitemguids == null) throw new ArgumentNullException();

            string[] splitEntries = weblistitemguids.SplitByChars("\n,");

            foreach (string line in splitEntries)
            {
                string trio = line.Trim();
                if (trio == "") continue;

                string[] info = trio.Split('#');
                WebListId webListId = new WebListId(info[0]);

                Guid webGuid = webListId.WebGuid;
                Guid listGuid = webListId.ListGuid;
                List<string> itemIdSstr = new List<string>(info[1].Split(';'));

                List<int> itemIds = new List<int>();
                foreach (string itemId in itemIdSstr)
                {
                    try
                    {
                        int id = int.Parse(itemId);
                        if (id > 0)
                        {
                            itemIds.Add(id);
                        }
                    }
                    // ReSharper disable once EmptyGeneralCatchClause
                    catch {}
                }

                Add(webGuid, listGuid, itemIds);
            }
        }

        #endregion

        #region ADD functions

        public void Add(Guid web, Guid list, int id)
        {
            if (web.IsEmpty()) throw new ArgumentValidationException("web", "is empty Guid");
            if (list.IsEmpty()) throw new ArgumentValidationException("list", "is empty Guid");

            if (!ContainsKey(web))
            {
                this[web] = new Dictionary<Guid, List<int>>();
            }

            if (!this[web].ContainsKey(list))
            {
                this[web][list] = new List<int>();
            }

            if (!this[web][list].Contains(id))
            {
                if (id > 0)
                {
                    this[web][list].Add(id);
                }
            }
        }

        public void Add(Guid web, Guid list, IEnumerable<int> ids)
        {
            if (web.IsEmpty()) throw new ArgumentValidationException("web", "is empty Guid");
            if (list.IsEmpty()) throw new ArgumentValidationException("list", "is empty Guid");

            foreach (int id in ids)
            {
                Add(web, list, id);
            }
        }

        public void Add(WebListItemId id)
        {
            if (id == null) throw new ArgumentNullException();

            Add(id.WebGuid, id.ListGuid, id.Item);
        }

        public void AddRange(IEnumerable<WebListItemId> ids)
        {
            if (ids == null) throw new ArgumentNullException();

            foreach (WebListItemId id in ids)
            {
                Add(id);
            }
        }

        public void AddRange(WebListItemDictionary webListItemDictionary)
        {
            if (webListItemDictionary == null) throw new ArgumentNullException();

            foreach (Guid webGuid in webListItemDictionary.Keys)
            {
                foreach (Guid listGuid in webListItemDictionary[webGuid].Keys)
                {
                    Add(webGuid, listGuid, webListItemDictionary[webGuid][listGuid]);
                }
            }
        }

        #endregion

        #region REMOVE functions

        //TODO

        #endregion

        #region PROCESS functions

        /// <summary>
        ///     Executes function on all items in a given list
        /// </summary>
        /// <param name="site"></param>
        /// <param name="func"></param>
        /// <returns></returns>
        public ICollection<object> ProcessItems(SPSite site, Func<SPListItem, object> func)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");

            using (new SPMonitoredScope("ProcessItems"))
            {
                List<object> results = new List<object>();

                foreach (KeyValuePair<Guid, Dictionary<Guid, List<int>>> kvpWeb in this)
                {
                    try
                    {
                        using (SPWeb web = site.OpenW(kvpWeb.Key, true))
                        {
                            foreach (KeyValuePair<Guid, List<int>> kvpList in kvpWeb.Value)
                            {
                                SPList list;

                                try
                                {
                                    list = web.OpenList(kvpList.Key, true);
                                }
                                catch (Exception listExc)
                                {
                                    results.Add(listExc);
                                    continue;
                                }

                                foreach (int itemId in kvpList.Value)
                                {
                                    SPListItem item;

                                    try
                                    {
                                        item = list.GetItemById(itemId);
                                    }
                                    catch (Exception)
                                    {
                                        results.Add(new SPListItemNotFoundException(itemId, list));
                                        continue;
                                    }

                                    results.Add(item.ProcessItem(func));
                                }
                            }
                        }
                    }
                    catch (Exception webExc)
                    {
                        results.Add(webExc);
                    }
                }

                return results;
            }
        }

        /// <summary>
        ///     Executes function on all items in a given list
        /// </summary>
        /// <param name="site"></param>
        /// <param name="versionValidAt">DateTime (in UTC!), which will be used to find the version that was valid at this point</param>
        /// <param name="func"></param>
        /// <returns></returns>
        public ICollection<object> ProcessItemVersion(SPSite site, DateTime versionValidAt,
            Func<SPListItemVersion, object> func)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");

            using (new SPMonitoredScope("ProcessItemVersion"))
            {
                List<object> results = new List<object>();

                foreach (KeyValuePair<Guid, Dictionary<Guid, List<int>>> kvpWeb in this)
                {
                    try
                    {
                        using (SPWeb web = site.OpenW(kvpWeb.Key, true))
                        {
                            foreach (KeyValuePair<Guid, List<int>> kvpList in kvpWeb.Value)
                            {
                                SPList list;

                                try
                                {
                                    list = web.OpenList(kvpList.Key, true);
                                }
                                catch (Exception listExc)
                                {
                                    results.Add(listExc);
                                    continue;
                                }

                                foreach (int itemId in kvpList.Value)
                                {
                                    SPListItem item;
                                    SPListItemVersion itemVersion = null;
                                    try
                                    {
                                        item = list.GetItemById(itemId);
                                        // FIND VERSION VALID AT HISTORIC DATE

                                        for (int i = 0; i < item.Versions.Count; i++)
                                        {
                                            itemVersion = item.Versions[i];

                                            // timezone
                                            var versionDate = itemVersion.Created;

                                            // posledni verze pred zacatkem pristi porady                            
                                            if (versionDate < versionValidAt)
                                            {
                                                break;
                                            }
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        results.Add(new SPListItemNotFoundException(itemId, list));
                                        continue;
                                    }

                                    results.Add(itemVersion.ProcessItem(func));
                                }
                            }
                        }
                    }
                    catch (Exception webExc)
                    {
                        results.Add(webExc);
                    }
                }

                return results;
            }
        }

        /// <summary>
        ///     Executes function under elevated rights on items of a given WebListItemDictionary structure
        /// </summary>
        /// <param name="site"></param>
        /// <param name="func"></param>
        /// <returns></returns>
        public ICollection<object> ProcessItemsElevated(SPSite site, Func<SPListItem, object> func)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");

            return (ICollection<object>) site.RunElevated(elevSite => ProcessItems(elevSite, func));
        }

        #endregion
    }

}