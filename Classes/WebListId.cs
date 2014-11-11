using System;
using System.Collections.Generic;
using System.Web;
using System.Web.Caching;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Navertica.SharePoint.Extensions;

// ReSharper disable once CheckNamespace
namespace Navertica.SharePoint
{
    /// <summary>
    ///     For passing and handling SPRequest-independent identification of lists
    /// </summary>
    public class WebListId
    {
        public Guid WebGuid;
        public Guid ListGuid;
        public string InvalidMessage;

        public bool IsValid
        {
            get { return !(WebGuid == Guid.Empty || ListGuid == Guid.Empty); }
        }

        #region Constructors

        public WebListId()
        {
        }

        public WebListId(Guid web, Guid list)
        {
            if (web.IsEmpty()) throw new ArgumentValidationException("web", "is empty Guid");
            if (list.IsEmpty()) throw new ArgumentValidationException("list", "is empty Guid");

            WebGuid = web;
            ListGuid = list;
        }

        /// <summary>
        ///     Returns WebListId for string in format "webguid:listguid"
        /// </summary>
        /// <param name="webListId"></param>
        public WebListId(string webListId)
        {
            if (webListId == null) throw new ArgumentNullException();

            const string cacheName = "WebListIdDictionary";
            Dictionary<string, WebListId> webListIdDict =
                (((Dictionary<string, WebListId>) HttpRuntime.Cache.Get(cacheName)) ??
                 new Dictionary<string, WebListId>());

            if (webListIdDict.ContainsKey(webListId))
            {
                WebGuid = webListIdDict[webListId].WebGuid;
                ListGuid = webListIdDict[webListId].ListGuid;
            }
            else
            {
                if (!webListId.Contains("http://") && !webListId.Contains("https://"))
                {
                    string[] webListGuids = webListId.Split(':');
                    try
                    {
                        WebGuid = new Guid(webListGuids[0]);
                        ListGuid = new Guid(webListGuids[1]);
                    }
                    catch
                    {
                        WebGuid = Guid.Empty;
                        ListGuid = Guid.Empty;
                        InvalidMessage = "Could not load WebListId from '" + webListId + "'";
                    }
                }
                else
                {
                    try
                    {
                        using (SPSite site = new SPSite(webListId))
                        {
                            site.RunElevated(
                                delegate(SPSite elevSite)
                                    //pod zvysenymi at to do cache neulozi clovek kterej nema prava ten web/list otevrit
                                {
                                    using (SPWeb web = elevSite.OpenW(webListId, true))
                                        //otevirat web muze jen z relativni URL
                                    {
                                        SPList list = web.OpenList(webListId);
                                        if (list != null)
                                        {
                                            WebGuid = web.ID;
                                            ListGuid = list.ID;
                                        }
                                        else
                                        {
                                            InvalidMessage = "Could not find list specified by '" + webListId +
                                                             "' - maybe was deleted or moved or user has no permissons!";
                                        }
                                    }

                                    return null;
                                });
                        }
                    }
                    catch
                    {
                        InvalidMessage = "Could not load WebListId from '" + webListId + "'";
                    }
                }

                if (IsValid)
                {
                    webListIdDict.Add(webListId, new WebListId(WebGuid, ListGuid));
                    HttpRuntime.Cache.Insert(cacheName, webListIdDict, null, DateTime.Now.AddDays(1),
                        Cache.NoSlidingExpiration, CacheItemPriority.Normal, null);
                }
            }
        }

        public WebListId(string webGuid, string listGuid)
        {
            if (webGuid == null) throw new ArgumentNullException("webGuid");
            if (listGuid == null) throw new ArgumentNullException("listGuid");

            try
            {
                WebGuid = new Guid(webGuid);
                ListGuid = new Guid(listGuid);
            }
            catch
            {
                WebGuid = Guid.Empty;
                ListGuid = Guid.Empty;
                InvalidMessage = "Could not load WebListId from '" + webGuid + "' and '" + listGuid + "'";
            }
        }

        public WebListId(SPList list)
        {
            if (list == null) throw new ArgumentNullException();

            WebGuid = list.ParentWeb.ID;
            ListGuid = list.ID;
        }

        #endregion

        #region Process Functions

        /// <summary>
        ///     Executes the function on all lists passed in listsForWeb dict.
        /// </summary>
        /// <param name="site"></param>
        /// <param name="func"></param>
        /// <returns>Result of delegate</returns>
        public object ProcessList(SPSite site, Func<SPList, object> func)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");
            if (!IsValid) throw new SPException("this instance of object is not valid: " + InvalidMessage);

            object result;

            using (new SPMonitoredScope("WebListId ProcessList"))
            {
                try
                {
                    if (site.RootWeb.ID == WebGuid)
                    {
                        SPList list = site.RootWeb.OpenList(ListGuid, true);
                        result = list.ProcessList(func);
                    }
                    else
                    {
                        using (SPWeb web = site.OpenW(WebGuid, true))
                        {
                            SPList list = web.OpenList(ListGuid, true);
                            result = list.ProcessList(func);
                        }
                    }
                }
                catch (Exception exc)
                {
                    result = exc;
                }
            }

            return result;
        }

        public object ProcessListElevated(SPSite site, Func<SPList, object> func)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");
            if (!IsValid) throw new SPException("this instance of object is not valid: " + InvalidMessage);

            return site.RunElevated(elevatedSite => ProcessList(elevatedSite, func));
        }

        /// <summary>
        ///     Executes function on items returned by given query string in a list given by WebListId structure
        /// </summary>
        /// <param name="site"></param>
        /// <param name="que"></param>
        /// <param name="func"></param>
        /// <param name="rowLimit"> </param>
        /// <returns></returns>
        /// <exception cref="SPWebAccesDeniedException"></exception>
        /// <exception cref="SPWebNotFoundException"></exception>
        /// <exception cref="SPListAccesDeniedException"></exception>
        /// <exception cref="SPListNotFoundException"></exception>
        public ICollection<object> ProcessItems(SPSite site, Func<SPListItem, object> func, Q que = null,
            int rowLimit = -1)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");
            if (!IsValid) throw new SPException("this instance of object is not valid: " + InvalidMessage);

            return ProcessItems(site, func, (que != null ? que.ToString() : ""), rowLimit);
        }

        /// <summary>
        ///     Executes function on items returned by given query string in a list given by WebListId structure
        /// </summary>
        /// <param name="site"></param>
        /// <param name="querystr"></param>
        /// <param name="func"></param>
        /// <param name="rowLimit"> </param>
        /// <returns></returns>
        /// <exception cref="SPWebAccesDeniedException"></exception>
        /// <exception cref="SPWebNotFoundException"></exception>
        /// <exception cref="SPListAccesDeniedException"></exception>
        /// <exception cref="SPListNotFoundException"></exception>
        public ICollection<object> ProcessItems(SPSite site, Func<SPListItem, object> func, string querystr = "",
            int rowLimit = -1)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");
            if (!IsValid) throw new SPException("this instance of object is not valid: " + InvalidMessage);

            ICollection<object> result = null;

            ProcessList(site, delegate(SPList list)
            {
                result = list.ProcessItems(func, querystr, rowLimit);
                return null;
            });

            return result;
        }

        /// <summary>
        ///     Executes function on items returned by given query string in a list given by WebListId structure
        /// </summary>
        /// <param name="site"></param>
        /// <param name="func"></param>
        /// <param name="que"> </param>
        /// <param name="rowLimit"> </param>
        /// <returns></returns>
        /// <exception cref="SPWebNotFoundException"></exception>
        /// <exception cref="SPListNotFoundException"></exception>
        public ICollection<object> ProcessItemsElevated(SPSite site, Func<SPListItem, object> func, Q que = null,
            int rowLimit = -1)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");
            if (!IsValid) throw new SPException("this instance of object is not valid: " + InvalidMessage);

            return ProcessItemsElevated(site, func, (que != null ? que.ToString() : ""), rowLimit);
        }

        /// <summary>
        ///     Executes function on items returned by given query string in a list given by WebListId structure
        /// </summary>
        /// <param name="site"></param>
        /// <param name="querystr"></param>
        /// <param name="func"></param>
        /// <param name="rowLimit"> </param>
        /// <returns></returns>
        /// <exception cref="SPWebNotFoundException"></exception>
        /// <exception cref="SPListNotFoundException"></exception>
        public ICollection<object> ProcessItemsElevated(SPSite site, Func<SPListItem, object> func, string querystr = "",
            int rowLimit = -1)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");
            if (!IsValid) throw new SPException("this instance of object is not valid: " + InvalidMessage);

            return
                (ICollection<object>)
                    site.RunElevated(elevatedSite => ProcessItems(elevatedSite, func, querystr, rowLimit));
        }

        #endregion

        public SPWeb OpenWeb(SPSite site, bool throwException = false)
        {
            return site.OpenW(WebGuid, throwException);
        }

        public SPList OpenList(SPWeb web, bool throwException = false)
        {
            return web.OpenList(ListGuid, throwException);
        }

        public override string ToString()
        {
            return IsValid
                ? string.Format("WebListId - WebGuid: {0}, ListGuid: {1}", WebGuid, ListGuid)
                : InvalidMessage;
        }
    }


}