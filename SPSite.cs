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
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Caching;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;

namespace Navertica.SharePoint.Extensions
{
    public static class SPSiteExtensions
    {
        /// <summary>
        /// Returns first positive list identification for a given list name or guid, starting on rootweb and then recursively walking the subwebs
        /// under elevated privileges (if possible - not sandboxed).
        /// </summary>
        /// <param name="site"></param>
        /// <param name="listNameOrGuid"></param>
        /// <returns>WebListId</returns>
        public static WebListId FindList(this SPSite site, string listNameOrGuid)
        {
            using (new SPMonitoredScope("FindList: " + listNameOrGuid))
            {
                if (site == null) throw new ArgumentNullException("site");
                if (listNameOrGuid == null) throw new ArgumentNullException("listNameOrGuid");
                if (listNameOrGuid == string.Empty) return new WebListId();

                bool set = false;
                string cacheName = "FindListSite_" + site.ID + "_" + listNameOrGuid;
                WebListId result = (WebListId) HttpRuntime.Cache.Get(cacheName);

                if (result != null && !site.InTimerJob())
                {
                    return new WebListId(result.WebGuid, result.ListGuid);
                }

                SPList list;

                result = site.RootWeb.FindList(listNameOrGuid);
                if (result != null && result.IsValid) set = true;

                if (!set)
                {
                    site.RunElevated(delegate(SPSite elevatedSite)
                    {
                        for (int i = 0; i < elevatedSite.AllWebs.Count; i++)
                        {
                            try
                            {
                                using (SPWeb otherWeb = elevatedSite.AllWebs[i])
                                {
                                    list = otherWeb.OpenList(listNameOrGuid);
                                    if (list != null)
                                    {
                                        result = new WebListId(list);
                                        return null;
                                    }
                                }
                            }
                            // ReSharper disable once EmptyGeneralCatchClause
                            catch (Exception) {}
                        }

                        return null;
                    });

                    if (result.IsValid)
                    {
                        HttpRuntime.Cache.Insert(cacheName, result, null, DateTime.Now.AddDays(1), Cache.NoSlidingExpiration, CacheItemPriority.Normal, null);
                    }
                }

                return result.IsValid
                    ? new WebListId(result.WebGuid, result.ListGuid)
                    : new WebListId { InvalidMessage = "Could not found list '" + listNameOrGuid + "' at site: " + site.Url };
            }
        }

        /// <summary>
        /// Finds all lists of given listType in given site and its webs
        /// </summary>
        /// <param name="site"></param>
        /// <param name="listType"></param>
        /// <returns></returns>
        public static WebListDictionary FindListsWithBaseTemplate(this SPSite site, SPListTemplateType listType)
        {
            if (site == null) throw new ArgumentNullException("site");

            using (new SPMonitoredScope("FindListsWithBaseTemplate at site " + site.Url + " listType: " + listType))
            {
                WebListDictionary foundLists = new WebListDictionary();

                site.RunElevated(delegate(SPSite elevatedSite)
                {
                    for (int i = 0; i < elevatedSite.AllWebs.Count; i++)
                    {
                        using (SPWeb web = elevatedSite.AllWebs[i])
                        {
                            foundLists.AddRange(web.Lists.Cast<SPList>().Where(l => l.BaseTemplate == listType).Select(l => new WebListId(l)));
                        }
                    }

                    return null;
                });

                return foundLists;
            }
        }

        /// <summary>
        /// Finds all lists containing content types with ids beginning with given ContentTypeId in given site and its webs
        /// </summary>
        /// <param name="site"></param>
        /// <param name="contentTypeId">initial section of a content type id - "0x0108" etc.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static WebListDictionary FindListsWithContentType(this SPSite site, string contentTypeId)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (contentTypeId == null) throw new ArgumentNullException("contentTypeId");

            SPContentTypeId loadedCtId;
            try
            {
                loadedCtId = new SPContentTypeId(contentTypeId);
            }
            catch (ArgumentException)
            {
                throw new ArgumentException("contentTypeId '" + contentTypeId + "' is not valid");
            }

            return FindListsWithContentType(site, loadedCtId);
        }

        /// <summary>
        /// Finds all lists containing content types with ids beginning with given ContentTypeId in given site and its webs
        /// </summary>
        /// <param name="site"></param>
        /// <param name="contentTypeId">initial section of a content type id - "0x0108" etc.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static WebListDictionary FindListsWithContentType(this SPSite site, SPContentTypeId contentTypeId)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (contentTypeId == null) throw new ArgumentNullException("contentTypeId");

            WebListDictionary foundLists = new WebListDictionary();

            using (new SPMonitoredScope("FindListsWithContentType at site " + site.Url + " contentTypeId: " + contentTypeId))
            {
                site.RunElevated(delegate(SPSite elevatedSite)
                {
                    for (int i = 0; i < elevatedSite.AllWebs.Count; i++)
                    {
                        using (SPWeb otherWeb = elevatedSite.AllWebs[i])
                        {
                            foreach (SPList list in otherWeb.Lists)
                            {
                                try
                                {
                                    if (list.ContentTypes.BestMatch(contentTypeId).IsChildOf(contentTypeId)) //if contentTypeId is "0x0108" it will add the children too
                                    {
                                        foundLists.Add(list);
                                    }
                                }
                                // ReSharper disable once EmptyGeneralCatchClause
                                catch {}
                            }
                        }
                    }
                    return null;
                });
            }

            return foundLists;
        }

        /// <summary>
        /// Returns WebListDictionary of all lists containing field with specified internal name
        /// </summary>
        /// <param name="site"></param>
        /// <param name="intFieldName"></param>
        /// <returns></returns>
        public static WebListDictionary FindListsWithField(this SPSite site, string intFieldName)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (intFieldName == null) throw new ArgumentNullException("intFieldName");

            WebListDictionary foundLists = new WebListDictionary();

            site.RunElevated(delegate(SPSite elevatedSite)
            {
                for (int i = 0; i < elevatedSite.AllWebs.Count; i++)
                {
                    using (SPWeb otherWeb = elevatedSite.AllWebs[i])
                    {
                        foreach (SPList list in otherWeb.Lists)
                        {
                            try
                            {
                                if (list.ContainsFieldIntName(intFieldName))
                                {
                                    foundLists.Add(list);
                                }
                            }
                            // ReSharper disable once EmptyGeneralCatchClause
                            catch {}
                        }
                    }
                }

                return null;
            });

            return foundLists;
        }

        /// <summary>
        /// Returns WebListItemId for both absolute and relative url of a document (not a list item)
        /// </summary>
        /// <param name="site"></param>
        /// <param name="url">a</param>
        /// <returns>WebListItemId</returns>
        /// <exception cref="SPWebNotFoundException"></exception>
        /// <exception cref="SPWebAccesDeniedException"></exception>
        public static WebListItemId GetItemByUrl(this SPSite site, string url)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (url == null) throw new ArgumentNullException("url");

            WebListItemId result = new WebListItemId();

            using (SPWeb web = site.OpenW(url, true))
            {
                try
                {
                    result = new WebListItemId(web.GetListItem(url));
                }
                catch (Exception e)
                {
                    result.InvalidMessage = e.ToString();
                }
            }

            return result;
        }

        /// <summary>
        /// Gets the admin token 
        /// http://dotnetfollower.com/wordpress/tag/runwithelevatedprivileges/
        /// </summary>
        /// <param name="site"></param>
        /// <returns></returns>
        public static SPUserToken GetSystemToken(this SPSite site)
        {
            if (site == null) throw new ArgumentNullException("site");

            SPUserToken res = null;
            bool originalCatchValue = SPSecurity.CatchAccessDeniedException;
            try
            {
                SPSecurity.CatchAccessDeniedException = false;
                res = site.SystemAccount.UserToken;
            }
            catch (UnauthorizedAccessException)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate
                {
                    using (SPSite elevatedSPSite = new SPSite(site.ID))
                    {
                        res = elevatedSPSite.SystemAccount.UserToken; // (***)
                    }
                });
            }
            finally
            {
                SPSecurity.CatchAccessDeniedException = originalCatchValue;
            }

            return res;
        }

        /// <summary>
        /// Checks if we're running in sandbox
        /// </summary>
        /// <param name="site"></param>
        /// <returns></returns>
        public static bool InSandbox(this SPSite site)
        {
            return AppDomain.CurrentDomain.FriendlyName.Contains("Sandbox");
        }

        /// <summary>
        /// Checks if we're running in OWSTIMER
        /// </summary>
        /// <param name="site"></param>
        /// <returns></returns>
        public static bool InTimerJob(this SPSite site)
        {
            return System.Diagnostics.Process.GetCurrentProcess().ProcessName.ToLowerInvariant() == "owstimer";
        }

        #region OpenW

        /// <summary>
        /// Return an open web identified by Guid - fallback solution for IronPython scripts in case there's a Guid instead of string. 
        /// Resulting SPWeb has to be DISPOSED MANUALLY
        /// </summary>
        /// <param name="site"></param>
        /// <param name="webGuid"></param>
        /// <param name="throwExc">should failure to open web throw Exception</param>
        /// <returns>open SPWeb or null</returns>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="SPWebAccesDeniedException"></exception>
        /// <exception cref="SPWebNotFoundException"></exception>
        public static SPWeb OpenW(this SPSite site, Guid webGuid, bool throwExc = false)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (webGuid == null) throw new ArgumentNullException("webGuid");
            if (webGuid.IsEmpty()) throw new ArgumentException("Guid is Empty");

            return OpenW(site, webGuid.ToString(), throwExc);
        }

        /// <summary>
        /// Return an open web identified by name, url or Guid in string.
        /// Resulting SPWeb has to be DISPOSED MANUALLY
        /// </summary>
        /// <param name="site"></param>
        /// <param name="webText"></param>
        /// <param name="throwExc">should failure to open web throw Exception</param>
        /// <param name="requireExactUrl">url http://portal/RiskPoint/default.aspx with true is null, with false will return SPWeb</param>
        /// <returns>open SPWeb or null</returns>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="SPWebAccesDeniedException"></exception>
        /// <exception cref="SPWebNotFoundException"></exception>
        public static SPWeb OpenW(this SPSite site, string webText, bool throwExc = false, bool requireExactUrl = false)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (webText == null) throw new ArgumentNullException("webText");

            using (new SPMonitoredScope("OpenW - " + webText))
            {
                bool originalCatchValue = SPSecurity.CatchAccessDeniedException;

                Guid webGuid = Guid.Empty;
                try
                {
                    webGuid = new Guid(webText);
                }
                // ReSharper disable once EmptyGeneralCatchClause
                catch {}

                if (!webGuid.IsEmpty())
                {
                    try
                    {
                        SPSecurity.CatchAccessDeniedException = false;
                        SPWeb webG = site.OpenWeb(webGuid);
                        SPSecurity.CatchAccessDeniedException = originalCatchValue;

                        if (webG.Exists) return webG;
                    }
                    catch (UnauthorizedAccessException)
                    {
                        if (throwExc) throw new SPWebAccesDeniedException(webGuid, site);
                    }
                    catch (FileNotFoundException)
                    {
                        if (throwExc) throw new SPWebNotFoundException(webGuid, site);
                    }
                    catch (DirectoryNotFoundException) //vyhazuje kdyz mam guid z jiny site
                    {
                        if (throwExc) throw new SPWebNotFoundException(webGuid, site);
                    }
                    catch
                    {
                        if (throwExc) throw;
                    }
                    return null;
                }

                webText = webText.Replace(site.Url, "").Trim();

                bool dispose = false;
                SPWeb web = null;

                try
                {
                    SPSecurity.CatchAccessDeniedException = false;
                    try
                    {
                        web = site.OpenWeb(webText, requireExactUrl);
                    }
                    catch (ArgumentException) 
                    {
                        if (webText.StartsWith("/"))
                        {
                            webText = webText.Remove(0, 1);
                            web = site.OpenWeb(webText, requireExactUrl);
                        }
                    }

                    if (web != null)
                    {
                        // ReSharper disable once UnusedVariable
                        string webTitle = web.Title; // web can be not null and at the same time throw either UnauthorizedAccessException or FileNotFoundException
                        if (web.Exists) return web;
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    dispose = true;
                    if (throwExc) throw new SPWebAccesDeniedException(webText, site);
                }
                catch (FileNotFoundException)
                {
                    dispose = true;
                    if (throwExc) throw new SPWebNotFoundException(webText, site);
                }
                catch
                {
                    dispose = true;
                    if (throwExc) throw;
                }
                finally
                {
                    if (dispose && web != null) web.Dispose();
                    SPSecurity.CatchAccessDeniedException = originalCatchValue;
                }
            }
            return null;
        }

        #endregion

        /// <summary>
        /// Run code as elevated admin
        /// </summary>
        /// <param name="site"></param>
        /// <param name="func"></param>
        /// <returns></returns>
        public static object RunElevated(this SPSite site, Func<SPSite, object> func)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (func == null) throw new ArgumentNullException("func");

            Guid siteGuid = site.ID;
            object result = null;


            if (site.InSandbox() || site.RootWeb.CurrentUser.IsSiteAdmin)
            {
                bool origSiteunsafe = site.AllowUnsafeUpdates;
                site.AllowUnsafeUpdates = true;
                result = func(site);
                site.AllowUnsafeUpdates = origSiteunsafe;
            }
            else
            {
                using (new SPMonitoredScope("RunElevated SPSite - " + site.Url))
                {
                    SPSecurity.RunWithElevatedPrivileges(delegate
                    {
                        using (SPSite elevatedSite = new SPSite(siteGuid, GetSystemToken(site)))
                        {
                            elevatedSite.AllowUnsafeUpdates = true;
                            result = func(elevatedSite);
                            elevatedSite.AllowUnsafeUpdates = false;
                        }
                    });
                }
            }

            return result;
        }

        /// <summary>
        /// Run code as custom user
        /// </summary>
        /// <param name="site"></param>
        /// <param name="user"> </param>
        /// <param name="func"></param>
        /// <returns></returns>
        public static object RunAsUser(this SPSite site, SPUser user, Func<SPSite, object> func)
        {
            if (site == null) throw new ArgumentNullException("site");
            if (user == null) throw new ArgumentNullException("user");
            if (func == null) throw new ArgumentNullException("func");

            Guid siteGuid = site.ID;
            object result;
            SPUserToken token = user.UserToken;

            if (site.InSandbox())
            {
                bool origSiteunsafe = site.AllowUnsafeUpdates;
                site.AllowUnsafeUpdates = true;
                result = func(site);
                site.AllowUnsafeUpdates = origSiteunsafe;
            }
            else
            {
                using (new SPMonitoredScope("SPSite - RunAsUser"))
                {
                    using (SPSite alternativeSite = new SPSite(siteGuid, token))
                    {
                        alternativeSite.AllowUnsafeUpdates = true;
                        result = func(alternativeSite);
                        alternativeSite.AllowUnsafeUpdates = false;
                    }
                }
            }

            return result;
        }
    }
}