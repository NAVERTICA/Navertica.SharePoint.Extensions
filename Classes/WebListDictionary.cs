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
using System.Collections.ObjectModel;
using System.Linq;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Navertica.SharePoint.Extensions;

// ReSharper disable once CheckNamespace
namespace Navertica.SharePoint
{
  
    /// <summary>
    ///     For passing SPRequest-independent identification of several lists across webs. Grouping the lists by web lets us
    ///     open the SPWeb just once for all the lists.
    /// </summary>
    public class WebListDictionary : Dictionary<Guid, List<Guid>>
    {
        public int ListGuidCount
        {
            get { return Keys.SelectMany(g => this[g]).Count(); }
        }

        #region Constructors

        public WebListDictionary()
        {
        }

        public WebListDictionary(string webListIds)
        {
            if (webListIds == null) throw new ArgumentNullException();

            Add(webListIds);
        }

        public WebListDictionary(WebListId id)
        {
            if (id == null) throw new ArgumentNullException();

            Add(id);
        }

        public WebListDictionary(IEnumerable<string> ids)
        {
            if (ids == null) throw new ArgumentNullException();

            foreach (string id in ids)
            {
                Add(id);
            }
        }

        public WebListDictionary(IEnumerable<WebListId> ids)
        {
            if (ids == null) throw new ArgumentNullException();

            foreach (WebListId id in ids)
            {
                Add(id);
            }
        }

        public WebListDictionary(WebListDictionary ids)
        {
            if (ids == null) throw new ArgumentNullException();

            AddRange(ids);
        }

        #endregion

        #region ADD functions

        public void Add(SPList list)
        {
            if (list == null) throw new ArgumentNullException();

            Add(list.ParentWeb.ID, list.ID);
        }

        public void Add(Guid web, Guid list)
        {
            if (web.IsEmpty()) throw new ArgumentValidationException("web", "is empty Guid");
            if (list.IsEmpty()) throw new ArgumentValidationException("list", "is empty Guid");

            if (!ContainsKey(web))
            {
                this[web] = new List<Guid>();
            }

            if (!this[web].Contains(list))
            {
                this[web].Add(list);
            }
        }

        public void Add(WebListId id)
        {
            if (id == null) throw new ArgumentNullException();
            if (!id.IsValid) throw new SPException(id.InvalidMessage);

            Add(id.WebGuid, id.ListGuid);
        }

        /// <summary>
        ///     Add items from a given string of "WebGuid:ListGuid, ..." - separator of pairs can be comma, semicolon, newline...
        /// </summary>
        /// <param name="webListIds"></param>
        /// <returns></returns>
        public void Add(string webListIds)
        {
            if (webListIds == null) throw new ArgumentNullException();

            foreach (string row in webListIds.Replace(" ", "").SplitByCharsDefault())
            {
                Add(new WebListId(row));
            }
        }

        public void AddRange(string webListIds)
        {
            if (webListIds == null) throw new ArgumentNullException();

            Add(webListIds);
        }

        public void AddRange(WebListDictionary webListDictionary)
        {
            if (webListDictionary == null) throw new ArgumentNullException();

            foreach (Guid webGuid in webListDictionary.Keys)
            {
                foreach (Guid listGuid in webListDictionary[webGuid])
                {
                    Add(webGuid, listGuid);
                }
            }
        }

        public void AddRange(IEnumerable<string> webListIds)
        {
            if (webListIds == null) throw new ArgumentNullException();

            foreach (string weblistId in webListIds)
            {
                Add(new WebListId(weblistId));
            }
        }

        public void AddRange(IEnumerable<WebListId> webListIds)
        {
            if (webListIds == null) throw new ArgumentNullException();

            foreach (WebListId id in webListIds)
            {
                Add(id);
            }
        }

        #endregion

        #region REMOVE functions

        public void Remove(Guid web, Guid list)
        {
            if (web.IsEmpty()) throw new ArgumentValidationException("web", "is empty Guid");
            if (list.IsEmpty()) throw new ArgumentValidationException("list", "is empty Guid");

            if (ContainsKey(web))
            {
                this[web].Remove(list);

                if (this[web].Count == 0)
                {
                    Remove(web);
                }
            }
        }

        public void Remove(WebListId id)
        {
            if (id == null) throw new ArgumentNullException();

            Remove(id.WebGuid, id.ListGuid);
        }

        /// <summary>
        ///     Remove items from a given string of "WebGuid:ListGuid, ..." - separator of pairs can be comma, semicolon,
        ///     newline...
        /// </summary>
        /// <param name="webListIds"></param>
        /// <returns></returns>
        public void Remove(string webListIds)
        {
            if (webListIds == null) throw new ArgumentNullException();

            foreach (string row in webListIds.Replace(" ", "").SplitByCharsDefault())
            {
                Remove(new WebListId(row));
            }
        }

        public void RemoveRange(string webListIds)
        {
            if (webListIds == null) throw new ArgumentNullException();
            Remove(webListIds);
        }

        public void RemoveRange(WebListDictionary webListDictionary)
        {
            if (webListDictionary == null) throw new ArgumentNullException();

            foreach (Guid webGuid in webListDictionary.Keys)
            {
                foreach (Guid listGuid in webListDictionary[webGuid])
                {
                    Remove(webGuid, listGuid);
                }
            }
        }

        public void RemoveRange(IEnumerable<string> weblistIds)
        {
            if (weblistIds == null) throw new ArgumentNullException();

            foreach (string weblistId in weblistIds)
            {
                Remove(new WebListId(weblistId));
            }
        }

        public void RemoveRange(IEnumerable<WebListId> weblistIds)
        {
            if (weblistIds == null) throw new ArgumentNullException();

            foreach (WebListId id in weblistIds)
            {
                Remove(id);
            }
        }

        #endregion

        #region PROCESS functions

        /// <summary>
        ///     Executes function on all webs of a given WebListDictionary structure
        /// </summary>
        /// <param name="site"></param>
        /// <param name="func"></param>
        /// <returns></returns>
        public ICollection<object> ProcessWebs(SPSite site, Func<SPWeb, object> func)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");

            ICollection<object> results = new Collection<object>();
            using (new SPMonitoredScope("WebListDictionary ProcessWebs"))
            {
                foreach (Guid webGuid in Keys)
                {
                    try
                    {
                        using (SPWeb web = site.OpenW(webGuid, true))
                        {
                            web.AllowUnsafeUpdates = true;

                            try
                            {
                                results.Add(func(web));
                            }
                            catch (Exception exc)
                            {
                                results.Add(exc);
                            }
                        }
                    }

                    catch (Exception exc)
                    {
                        results.Add(exc);
                    }
                }
            }

            return results;
        }

        /// <summary>
        ///     Executes function on all webs of a given WebListDictionary structure with Elevated Privileges
        /// </summary>
        /// <param name="site"></param>
        /// <param name="func"></param>
        /// <returns></returns>
        public ICollection<object> ProcessWebsElevated(SPSite site, Func<SPWeb, object> func)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");

            return (ICollection<object>) site.RunElevated(elevatedSite => ProcessWebs(elevatedSite, func));
        }

        /// <summary>
        ///     Executes the function on all lists passed in listsForWeb dict.
        /// </summary>
        /// <param name="site"></param>
        /// <param name="func"></param>
        /// <returns>Result of delegate</returns>
        public ICollection<object> ProcessLists(SPSite site, Func<SPList, object> func)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");

            ICollection<object> results = new Collection<object>();

            using (new SPMonitoredScope("WebListDictionary ProcessLists - Count:" + ListGuidCount))
            {
                foreach (Guid webGuid in Keys)
                {
                    try
                    {
                        using (SPWeb web = site.OpenW(webGuid, true))
                        {
                            foreach (Guid listGuid in this[webGuid])
                            {
                                using (new SPMonitoredScope("ProcessList"))
                                {
                                    SPList list;
                                    try
                                    {
                                        list = web.OpenList(listGuid, true);
                                    }
                                    catch (Exception liExc)
                                    {
                                        results.Add(liExc);
                                        continue;
                                    }

                                    try
                                    {
                                        results.Add(list.ProcessList(func));
                                    }
                                    catch (Exception exc)
                                    {
                                        results.Add(exc);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception exc)
                    {
                        results.Add(exc);
                    }
                }
            }
            return results;
        }

        public ICollection<object> ProcessListsElevated(SPSite site, Func<SPList, object> func)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");

            return (ICollection<object>) site.RunElevated(elevatedSite => ProcessLists(elevatedSite, func));
        }

        public ICollection<object> ProcessItems(SPSite site, Func<SPListItem, object> func, Q que = null,
            int rowLimit = -1)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");

            return ProcessItems(site, func, (que != null ? que.ToString() : ""), rowLimit);
        }

        /// <summary>
        ///     Executes function on all items in all lists of a given WebListDictionary structure
        /// </summary>
        /// <param name="site"></param>
        /// <param name="querystr"></param>
        /// <param name="func"></param>
        /// <param name="rowLimit"></param>
        /// <returns></returns>
        public ICollection<object> ProcessItems(SPSite site, Func<SPListItem, object> func, string querystr = "",
            int rowLimit = -1)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");

            List<object> result = new List<object>();

            foreach (KeyValuePair<Guid, List<Guid>> kvp in this)
            {
                try
                {
                    using (SPWeb web = site.OpenW(kvp.Key, true))
                    {
                        foreach (Guid listGuid in kvp.Value)
                        {
                            SPList list;
                            try
                            {
                                list = web.OpenList(listGuid, true);
                            }
                            catch (Exception listExc)
                            {
                                result.Add(listExc);
                                continue;
                            }

                            result.AddRange(list.ProcessItems(func, querystr, rowLimit));
                        }
                    }
                }
                catch (Exception webExc)
                {
                    result.Add(webExc);
                }
            }

            return result;
        }

        /// <summary>
        ///     Executes function on all items in all lists of a given WebListDictionary structure
        /// </summary>
        /// <param name="site"></param>
        /// <param name="que"></param>
        /// <param name="func"></param>
        /// <param name="rowLimit"></param>
        /// <returns></returns>
        public ICollection<object> ProcessItemsElevated(SPSite site, Func<SPListItem, object> func, Q que = null,
            int rowLimit = -1)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");

            return
                (ICollection<object>)
                    site.RunElevated(
                        elevatedSite => ProcessItems(elevatedSite, func, (que != null ? que.ToString() : ""), rowLimit));
        }

        /// <summary>
        ///     Executes function on all items in all lists of a given WebListDictionary structure
        /// </summary>
        /// <param name="site"></param>
        /// <param name="querystr"></param>
        /// <param name="func"></param>
        /// <param name="rowLimit"></param>
        /// <returns></returns>
        public ICollection<object> ProcessItemsElevated(SPSite site, Func<SPListItem, object> func, string querystr = "",
            int rowLimit = -1)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");

            return
                (ICollection<object>)
                    site.RunElevated(elevatedSite => ProcessItems(elevatedSite, func, querystr, rowLimit));
        }

        #endregion

        public List<string> GetListUrls(SPSite site)
        {
            if (site == null) throw new ArgumentNullException();

            List<string> result = new List<string>();
            ProcessLists(site, delegate(SPList list)
            {
                result.Add(list.AbsoluteUrl());
                return null;
            });
            return result;
        }

        public bool Contains(SPList list)
        {
            if (list == null) throw new ArgumentNullException();

            return ContainsKey(list.ParentWeb.ID) && this[list.ParentWeb.ID].Contains(list.ID);
        }

        public string ToString(SPSite site)
        {
            string res = "";
            ProcessLists(site, delegate(SPList list)
            {
                res += list.AbsoluteUrl() + "\n";
                return null;
            });
            return res;
        }
    }

}