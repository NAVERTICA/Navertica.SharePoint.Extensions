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
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Navertica.SharePoint.Extensions;

// ReSharper disable once CheckNamespace
namespace Navertica.SharePoint
{
   
    /// <summary>
    ///     For passing and handling SPRequest-independent identification of list items
    /// </summary>
    public class WebListItemId
    {
        public Guid WebGuid;
        public Guid ListGuid;
        public int Item;
        public string InvalidMessage;

        public bool IsValid
        {
            get { return !(WebGuid == Guid.Empty || ListGuid == Guid.Empty || Item <= 0); }
        }

        #region Constructors

        public WebListItemId()
        {
        }

        /// <summary>
        ///     Listurl#ItemId or webGuid:listGuid:ItemId
        /// </summary>
        /// <param name="webListItemId"></param>
        public WebListItemId(string webListItemId)
        {
            if (webListItemId == null) throw new ArgumentNullException();

            string invalidMessage = "Could not load WebListItemId from '" + webListItemId + "'";

            if (webListItemId.Contains("#"))
            {
                string[] info = webListItemId.Split('#');

                if (info.Length == 2)
                {
                    WebListId webListId = new WebListId(info[0]);
                    if (webListId.IsValid)
                    {
                        int itemId;
                        Int32.TryParse(info[1], out itemId);
                        if (itemId > 0)
                        {
                            WebGuid = webListId.WebGuid;
                            ListGuid = webListId.ListGuid;
                            Item = itemId;
                        }
                        else InvalidMessage = webListId.InvalidMessage;
                    }
                    else InvalidMessage = webListId.InvalidMessage;
                }
                else InvalidMessage = invalidMessage;
            }
            else if (webListItemId.Contains("http://") && !webListItemId.Contains("https://"))
            {
                WebListId webListId = new WebListId(webListItemId);
                if (webListId.IsValid)
                {
                    try
                    {
                        //nacteme jak urlValue protoze muzeme predat string ve formatu "http://portal13devel/Lists/TEST/DisplayForm.aspx?ID=1, TEST"
                        SPFieldUrlValue urlValue = new SPFieldUrlValue(webListItemId);
                        webListItemId = urlValue.Url;
                    }
                    // ReSharper disable once EmptyGeneralCatchClause
                    catch {}

                    object id = webListItemId.GetParametersFromUrl().TryGetValue("id");

                    if (id == null) InvalidMessage = "Can not get parameter id from '" + webListItemId + "'";
                    else
                    {
                        int itemId;
                        Int32.TryParse(id.ToString(), out itemId);
                        if (itemId > 0)
                        {
                            WebGuid = webListId.WebGuid;
                            ListGuid = webListId.ListGuid;
                            Item = itemId;
                        }
                        else InvalidMessage = invalidMessage;
                    }
                }
                else InvalidMessage = webListId.InvalidMessage;
            }
            else if (webListItemId.Contains(":"))
            {
                string[] info = webListItemId.Split(':');

                if (info.Length == 3)
                {
                    int itemId;
                    Int32.TryParse(info[2], out itemId);
                    if (itemId > 0)
                    {
                        try
                        {
                            WebGuid = new Guid(info[0]);
                            ListGuid = new Guid(info[1]);
                            Item = itemId;
                        }
                        catch
                        {
                            WebGuid = Guid.Empty;
                            ListGuid = Guid.Empty;
                            InvalidMessage = invalidMessage;
                        }
                    }
                    else InvalidMessage = invalidMessage;
                }
                else InvalidMessage = invalidMessage;
            }
            else InvalidMessage = invalidMessage;
        }

        public WebListItemId(Guid web, Guid list, int item)
        {
            if (web.IsEmpty()) throw new ArgumentValidationException("web", "is empty Guid");
            if (list.IsEmpty()) throw new ArgumentValidationException("list", "is empty Guid");
            if (item <= 0) throw new ArgumentValidationException("item", "could not be less or equal zero");

            WebGuid = web;
            ListGuid = list;
            Item = item;
        }

        public WebListItemId(string webGuid, string listGuid, int item)
        {
            if (webGuid == null) throw new ArgumentNullException("webGuid");
            if (listGuid == null) throw new ArgumentNullException("listGuid");
            if (item <= 0) throw new ArgumentValidationException("item", "could not be less or equal zero");

            try
            {
                WebGuid = new Guid(webGuid);
                ListGuid = new Guid(listGuid);
                Item = item;
            }
            catch
            {
                WebGuid = Guid.Empty;
                ListGuid = Guid.Empty;
                InvalidMessage = "Could not load WebListItemId from '" + webGuid + "' and '" + listGuid + "' and '" +
                                 item + "'";
            }
        }

        public WebListItemId(SPListItem item)
        {
            if (item == null) throw new ArgumentNullException();

            WebGuid = item.ParentList.ParentWeb.ID;
            ListGuid = item.ParentList.ID;
            Item = item.ID;
        }

        #endregion

        /// <summary>
        ///     Executes the delegate on the item. Returns whatever the delegate returns.
        /// </summary>
        /// <param name="site"></param>
        /// <param name="func"></param>
        /// <returns>result of delegate</returns>
        /// <exception cref="SPWebAccesDeniedException"></exception>
        /// <exception cref="SPWebNotFoundException"></exception>
        /// <exception cref="SPListAccesDeniedException"></exception>
        /// <exception cref="SPListNotFoundException"></exception>
        public object ProcessItem(SPSite site, Func<SPListItem, object> func)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");
            if (!IsValid) throw new SPException("this instance of object is not valid: " + InvalidMessage);

            using (new SPMonitoredScope("WebListItemId ProcessItem"))
            {
                object result;

                using (SPWeb web = site.OpenW(WebGuid, true))
                {
                    SPList list = web.OpenList(ListGuid, true);

                    SPListItem item;
                    try
                    {
                        item = list.GetItemById(Item);
                    }
                    catch
                    {
                        throw new SPListItemNotFoundException(Item, list);
                    }

                    result = item.ProcessItem(func);
                }
                return result;
            }
        }

        /// <summary>
        ///     Opens the item under elevated privilages and executes the delegate on the elevated item. Returns whatever the
        ///     delegate returns.
        /// </summary>
        /// <param name="site"></param>
        /// <param name="func"></param>
        /// <returns>result of delegate</returns>
        /// <exception cref="SPWebNotFoundException"></exception>
        /// <exception cref="SPListNotFoundException"></exception>
        public object RunElevated(SPSite site, Func<SPListItem, object> func)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");
            if (!IsValid) throw new SPException("this instance of object is not valid: " + InvalidMessage);

            return site.RunElevated(elevSite => ProcessItem(elevSite, func));
        }

        public SPWeb OpenWeb(SPSite site)
        {
            return site.OpenW(WebGuid);
        }

        public SPList OpenList(SPWeb web)
        {
            return web.OpenList(ListGuid);
        }

        public override string ToString()
        {
            return IsValid
                ? string.Format("WebListItemId - WebGuid: {0}, ListGuid: {1}, Item: {2}", WebGuid, ListGuid, Item)
                : InvalidMessage;
        }
    }

  

}