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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Caching;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;

namespace Navertica.SharePoint.Extensions
{
    public static class SPWebExtensions
    {
        /// <summary>
        /// Uploads new document
        /// </summary>
        /// <param name="web"></param>
        /// <param name="fullPath"></param>
        /// <param name="uploadStream"></param>
        /// <returns></returns>
        public static SPListItem CreateNewDocument(this SPWeb web, string fullPath, Stream uploadStream)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (fullPath == null) throw new ArgumentNullException("fullPath");
            if (uploadStream == null) throw new ArgumentNullException("uploadStream");

            WebRequest request = WebRequest.Create(fullPath);
            request.Credentials = CredentialCache.DefaultCredentials; // User must have 'Contributor' access to the document library
            request.Method = "PUT";
            request.Headers.Add("Overwrite", "t");

            byte[] buffer = new byte[4096];
            using (Stream stream = request.GetRequestStream())
            {
                for (int i = uploadStream.Read(buffer, 0, buffer.Length); i > 0; i = uploadStream.Read(buffer, 0, buffer.Length))
                {
                    stream.Write(buffer, 0, i);
                }
                stream.Close();
            }
            WebResponse response = request.GetResponse(); // Upload the file
            response.Close();

            return web.GetItemByUrl(fullPath);
        }

        /// <summary>
        /// Delete web with all subwebs
        /// </summary>
        /// <param name="web"></param>
        public static void DeleteAll(this SPWeb web)
        {
            if (web == null) throw new ArgumentNullException("web");

            for (int i = 0; i < web.Webs.Count; i++)
            {
                using (SPWeb subWeb = web.Webs[i])
                {
                    subWeb.DeleteAll();
                }
            }
            web.Delete();
        }

        #region FindList/s

        /// <summary>
        /// Recursive - returns first positive list identification for a given list name or guid, starting with web and then recursively walking its subwebs
        /// </summary>
        /// <param name="web"></param>
        /// <param name="listNameUrlOrGuid"> </param>
        /// <returns>WebListId or WebListId with empty Guids if not found</returns>
        public static WebListId FindList(this SPWeb web, string listNameUrlOrGuid) // TODO - prejmenovat na FindListRecursive
        {
            if (web == null) throw new ArgumentNullException("web");
            if (listNameUrlOrGuid == null) throw new ArgumentNullException("listNameUrlOrGuid");
            if (listNameUrlOrGuid == string.Empty) return new WebListId();

            string cacheName = "FindListWeb_" + web.ID + "_" + listNameUrlOrGuid;
            WebListId result = (WebListId) HttpRuntime.Cache.Get(cacheName);
            if (result != null && !web.Site.InTimerJob())
            {
                return new WebListId(result.WebGuid, result.ListGuid);
            }

            try
            {
                SPList list = web.OpenList(listNameUrlOrGuid);

                if (list == null)
                {
                    if (web.ParentWeb != null)
                    {
                        result = web.ParentWeb.FindList(listNameUrlOrGuid);
                    }
                }

                if (result == null && list != null)
                {
                    result = new WebListId(list);
                }
            }
            catch (FileNotFoundException) //nastane pri mazani webu
            {
                result = new WebListId();
            }

            if (result == null) result = new WebListId();

            if (result.IsValid)
            {
                HttpRuntime.Cache.Insert(cacheName, result, null, DateTime.Now.AddDays(1), Cache.NoSlidingExpiration, CacheItemPriority.Normal, null);
            }

            return result.IsValid
                ? new WebListId(result.WebGuid, result.ListGuid)
                : new WebListId { InvalidMessage = "Could not found list '" + listNameUrlOrGuid + "' at web: " + web.Url };
        }

        /// <summary>
        /// Finds all lists of given listType in given web and subwebs 
        /// </summary>
        /// <param name="web"></param>
        /// <param name="listType"></param>
        /// <returns></returns>
        public static WebListDictionary FindListsWithBaseTemplate(this SPWeb web, SPListTemplateType listType)
        {
            if (web == null) throw new ArgumentNullException("web");

            WebListDictionary foundLists = new WebListDictionary();
            using (new SPMonitoredScope("FindListsWithBaseTemplate at web " + web.Url + " listType: " + listType))
            {
                IEnumerable<Guid> webIDs = GetIds(web);

                web.Site.RunElevated(delegate(SPSite elevatedSite)
                {
                    foreach (Guid childWebId in webIDs)
                    {
                        using (SPWeb childWeb = elevatedSite.OpenW(childWebId, true))
                        {
                            foundLists.AddRange(childWeb.Lists.Cast<SPList>().Where(l => l.BaseTemplate == listType).Select(l => new WebListId(l)));
                        }
                    }

                    return null;
                });
            }

            return foundLists;
        }

        /// <summary>
        /// Finds all lists containing content types with ids beginning with given ContentTypeId in given web and its subwebs
        /// </summary>
        /// <param name="web"></param>
        /// <param name="contentTypeId">initial section of a content type id - "0x0108" etc.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static WebListDictionary FindListsWithContentType(this SPWeb web, string contentTypeId)
        {
            if (web == null) throw new ArgumentNullException("web");
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

            return FindListsWithContentType(web, loadedCtId);
        }

        /// <summary>
        /// Finds all lists containing content types with ids beginning with given ContentTypeId in given web and its subwebs
        /// </summary>
        /// <param name="web"></param>
        /// <param name="contentTypeId">initial section of a content type id - "0x0108" etc.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static WebListDictionary FindListsWithContentType(this SPWeb web, SPContentTypeId contentTypeId)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (contentTypeId == null) throw new ArgumentNullException("contentTypeId");

            WebListDictionary foundLists = new WebListDictionary();
            using (new SPMonitoredScope("FindListsWithContentType at web " + web.Url + " contentTypeId: " + contentTypeId))
            {
                IEnumerable<Guid> webIDs = GetIds(web);

                web.Site.RunElevated(delegate(SPSite elevatedSite)
                {
                    foreach (Guid childWebId in webIDs)
                    {
                        using (SPWeb childWeb = elevatedSite.OpenW(childWebId, true))
                        {
                            foreach (SPList list in childWeb.Lists)
                            {
                                try
                                {
                                    if (list.ContentTypes.BestMatch(contentTypeId).IsChildOf(contentTypeId)) //if ct is "0x0108" it will add the childs too
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

        public static WebListDictionary FindListsWithField(this SPWeb web, string intFieldName)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (intFieldName == null) throw new ArgumentNullException("intFieldName");

            WebListDictionary foundLists = new WebListDictionary();
            IEnumerable<Guid> webIDs = GetIds(web);

            web.Site.RunElevated(delegate(SPSite elevatedSite)
            {
                foreach (Guid childWebId in webIDs)
                {
                    using (SPWeb childWeb = elevatedSite.OpenW(childWebId, true))
                    {
                        foreach (SPList list in childWeb.Lists)
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
        /// Returns the list of Guid of the subwebs and currentWeb. Must be called with elevated privileges
        /// </summary>
        /// <param name="web"></param>
        /// <returns></returns>
        private static IEnumerable<Guid> GetIds(SPWeb web)
        {
            using (new SPMonitoredScope("GetWebIds for web " + web.Url))
            {
                List<Guid> webIDs = new List<Guid> { web.ID };

                web.RunElevated(delegate(SPWeb elevatedWeb)
                {
                    foreach (SPWeb subweb in elevatedWeb.Webs)
                    {
                        webIDs.Add(subweb.ID);
                        subweb.Dispose();
                    }

                    return null;
                });

                return webIDs;
            }
        }

        #endregion

        /// <summary>
        /// Tries to get SPFolder from given Url
        /// </summary>
        /// <param name="web"></param>
        /// <param name="folderUrl"></param>
        /// <returns>SPFolder or null</returns>
        public static SPFolder GetFolderFromUrl(this SPWeb web, string folderUrl)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (folderUrl == null) throw new ArgumentNullException("folderUrl");

            string resultFolder = folderUrl;

            if (folderUrl.EndsWith("Forms"))
            {
                resultFolder = folderUrl.Substring(0, folderUrl.LastIndexOf("Forms", StringComparison.InvariantCulture));
            }
            if (folderUrl.EndsWith("Forms/"))
            {
                resultFolder = folderUrl.Substring(0, folderUrl.LastIndexOf("Forms/", StringComparison.InvariantCulture));
            }

            SPFolder folder = web.GetFolder(resultFolder);

            return folder.Exists ? folder : null;
        }

        /// <summary>
        /// Returns SPListItem for both absolute and relative url of a document (not a list item)
        /// </summary>
        /// <param name="web"></param>
        /// <param name="url"></param>
        /// <returns></returns>
        public static SPListItem GetItemByUrl(this SPWeb web, string url) //TODO asi zbytecna funkce
        {
            if (web == null) throw new ArgumentNullException("web");
            if (url == null) throw new ArgumentNullException("url");

            return web.GetListItem(url);
        }

        #region Users Tools

        /// <summary>
        /// Get Principal's LoginName [lowercase format] by identification
        /// </summary>
        /// <param name="web"></param>
        /// <param name="identification">Can be PrincipalId, SPFieldUser, SPFieldUserValue(lookupformat), LoginName, Name or SPRoleAssigment</param>
        /// <returns>login or null</returns>
        public static string GetSPPrincipalLoginName(this SPWeb web, object identification)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (identification == null) throw new ArgumentNullException("identification");

            SPPrincipal principal = GetSPPrincipal(web, identification);
            return principal != null ? principal.LoginNameNormalized() : null;
        }

        /// <summary>
        /// Get Principal's Name by identification
        /// </summary>
        /// <param name="web"></param>
        /// <param name="identification">Can be PrincipalId, SPFieldUser, SPFieldUserValue(lookupformat), LoginName, Name or SPRoleAssigment</param>
        /// <returns>Name or null</returns>
        public static string GetSPPrincipalName(this SPWeb web, object identification)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (identification == null) throw new ArgumentNullException("identification");

            SPPrincipal principal = GetSPPrincipal(web, identification);
            return principal != null ? principal.Name : null;
        }

        /// <summary>
        /// Get SPUser by identification
        /// </summary>
        /// <param name="web"></param>
        /// <param name="identification">
        /// Can be UserId, SPFieldUser, SPFieldUserValue(lookupformat), LoginName, Name or SPRoleAssigment.
        /// If identification is SPFieldUserValueCollection or SPRoleAssignmentCollection returns null.
        /// </param>
        /// <returns>SPUser or null</returns>
        public static SPUser GetSPUser(this SPWeb web, object identification)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (identification == null) throw new ArgumentNullException("identification");
            string identAsString = identification.ToString().Trim();

            switch (identification.GetType().Name)
            {
                case "SPGroup":
                    return null;

                case "SPUser":
                    return (SPUser) identification;

                case "SPFieldUserValueCollection":
                    SPFieldUserValueCollection usrcol = (SPFieldUserValueCollection) identification;
                    return usrcol.Count == 1 ? GetSPUser(web, usrcol[0]) : null;

                case "SPRoleAssignmentCollection":
                    SPRoleAssignmentCollection rolecol = (SPRoleAssignmentCollection) identification;
                    return rolecol.Count == 1 ? GetSPUser(web, rolecol[0]) : null;

                case "SPRoleAssignment":
                    SPRoleAssignment roleAssignment = (SPRoleAssignment) identification;
                    return roleAssignment.Member is SPUser ? web.SiteUsers.GetByID(roleAssignment.Member.ID) : null;
            }

            try
            {
                return web.SiteUsers.GetByID(Convert.ToInt32(identification));
            }
                // ReSharper disable once EmptyGeneralCatchClause
            catch {}

            if (identAsString.IndexOf("#", StringComparison.InvariantCulture) > -1) // varianta s pouzitim lookup hodnoty
            {
                try
                {
                    return web.SiteUsers.GetByID(Convert.ToInt32(identAsString.GetLookupIndex()));
                }
                // ReSharper disable once EmptyGeneralCatchClause
                catch {}
            }

            try
            {
                if (identAsString.StartsWith("-1;#i")) identAsString = identAsString.Replace("-1;#i:", "i:");

                return web.SiteUsers[identAsString];
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch {}

            // resolve principal - both Login and Name
            try
            {
                SPPrincipalInfo principal = SPUtility.ResolvePrincipal(web, identAsString, SPPrincipalType.User, SPPrincipalSource.All, null, false);
                if (principal != null)
                {
                    if (principal.PrincipalId == -1)
                    {
                        SPUser user = null;
                        web.RunWithAllowUnsafeUpdates(delegate
                        {
                            user = web.EnsureUser(principal.LoginName);
                        });

                        return user;
                    }

                    return web.SiteUsers[principal.LoginName];
                }
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch {} // in case we get handed a string similar to claim, there's an exception

            //may be a domain group
            try
            {
                SPPrincipalInfo principal = SPUtility.ResolvePrincipal(web, identAsString, SPPrincipalType.SecurityGroup,
                    SPPrincipalSource.All, null, false);
                if (principal != null)
                {
                    if (principal.PrincipalId == -1)
                    {
                        SPUser user = null;
                        web.RunWithAllowUnsafeUpdates(delegate
                        {
                            user = web.EnsureUser(principal.LoginName);
                        });

                        return user;
                    }
                    return web.SiteUsers.GetByID(principal.PrincipalId);
                }
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch { } // in case we get handed a string similar to claim, there's an exception

            return null;
        }

        /// <summary>
        /// Get SPGroup by identification
        /// </summary>
        /// <param name="web"></param>
        /// <param name="identification">
        /// Can be GroupId, SPFieldUserValue(lookupformat), LoginName, Name or SPRoleAssigment.
        /// If identification is SPFieldUserValueCollection or SPRoleAssignmentCollection returns null.
        /// </param>
        /// <returns>SPGroup or null</returns>
        public static SPGroup GetSPGroup(this SPWeb web, object identification)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (identification == null) throw new ArgumentNullException("identification");

            switch (identification.GetType().Name)
            {
                case "SPGroup":
                    return (SPGroup) identification;

                case "SPUser":
                    return null;

                case "SPFieldUserValueCollection":
                case "SPRoleAssignmentCollection":
                    return null;
                case "SPRoleAssignment":
                {
                    SPRoleAssignment roleAssignment = (SPRoleAssignment) identification;
                    if (roleAssignment.Member is SPGroup)
                    {
                        try
                        {
                            return web.SiteGroups.GetByID(roleAssignment.Member.ID);
                        }
                        // ReSharper disable once EmptyGeneralCatchClause
                        catch (Exception) {}
                    }
                    else
                    {
                        return null;
                    }
                    break;
                }
            }

            try
            {
                return web.SiteGroups.GetByID(Convert.ToInt32(identification));
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch {}

            string identificationString = identification.ToString();

            if (identificationString.IndexOf("#", StringComparison.InvariantCulture) > -1) // varianta s pouzitim lookup hodnoty
            {
                try
                {
                    return web.SiteGroups.GetByID(Convert.ToInt32(identificationString.GetLookupIndex()));
                }
                // ReSharper disable once EmptyGeneralCatchClause
                catch {}
            }

            try
            {
                return web.SiteGroups[identificationString];
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch {}

            // varianta s resolve principal - resolvne jak login tak Name
            string groupStr = identificationString.GetLookupValue().Trim();
            SPPrincipalInfo principal = SPUtility.ResolvePrincipal(web, groupStr, SPPrincipalType.SharePointGroup, SPPrincipalSource.All, null, false);
            if (principal != null)
            {
                return web.SiteGroups.GetByID(principal.PrincipalId);
            }

            return null;
        }

        /// <summary>
        /// Get SPPrincipal(SPGroup or SPUser) by identification
        /// </summary>
        /// <param name="web"></param>
        /// <param name="identification">
        /// Can be UserId, GroupId, SPFieldUser, SPFieldUserValue(lookupformat), LoginName, Name or SPRoleAssigment.
        /// If identification is SPFieldUserValueCollection or SPRoleAssignmentCollection returns null
        /// </param>
        /// <returns>SPPrincipal or null</returns>
        public static SPPrincipal GetSPPrincipal(this SPWeb web, object identification)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (identification == null) throw new ArgumentNullException("identification");

            if (identification.GetType().Name.EqualAny(new[] { "SPFieldUserValueCollection", "SPRoleAssignmentCollection" })) return null;

            if (identification is SPPrincipal) return (SPPrincipal) identification;

            SPPrincipal principal = (SPPrincipal) GetSPGroup(web, identification) ?? GetSPUser(web, identification);

            return principal;
        }

        /// <summary>
        /// Get Principals defined by array of identifications
        /// </summary>
        /// <param name="web"></param>
        /// <param name="identification"></param>
        /// <returns>List of SPPrincipal, which are not duplicated and was susscesfully load. List does not contains null elements.</returns>
        public static List<SPPrincipal> GetSPPrincipals(this SPWeb web, object identification)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (identification == null) throw new ArgumentNullException("identification");

            return GetSPPrincipals(web, new List<object> { identification });
        }

        /// <summary>
        /// Get Principals defined by array of identifications or straight by string
        /// </summary>
        /// <param name="web"></param>
        /// <param name="identifications"></param>
        /// <returns>List of SPPrincipal, which are not duplicated and was susscesfully load. List does not contains null elements.</returns>
        public static List<SPPrincipal> GetSPPrincipals(this SPWeb web, IEnumerable identifications) //TODO prepsat
        {
            if (web == null) throw new ArgumentNullException("web");
            if (identifications == null) throw new ArgumentNullException("identifications");

            List<SPPrincipal> principals = new List<SPPrincipal>();
            using (new SPMonitoredScope("GetSPPrincipals"))
            {
                // ReSharper disable PossibleMultipleEnumeration
                if (identifications is string) return GetSPPrincipals(web, new[] { identifications }); //Muze byt primo string

                List<string> processLog = new List<string>();

                try
                {
                    foreach (object identification in identifications.Cast<object>().Where(i => i != null))
                    {
                        string identAsString = identification.ToString();
                        if (identification is SPRoleAssignment)
                        {
                            processLog.Add(( (SPRoleAssignment) identification ).Member.ToString());
                        }
                        else
                        {
                            processLog.Add(identification.ToString());
                        }

                        if (identification is SPRoleAssignmentCollection)
                        {
                            SPRoleAssignmentCollection col = (SPRoleAssignmentCollection) identification;
                            principals.AddRange(col.Cast<SPRoleAssignment>().Select(role => role.Member));
                        }
                        else if (identAsString.Contains(";#")) //Muze byt primo string
                        {
                            SPFieldUserValueCollection col = new SPFieldUserValueCollection(web, identification.ToString());
                            principals.AddRange(col.Select(userValue => GetSPPrincipal(web, userValue)));
                        }
                        else if (identAsString.Contains(";")) // melo by se pouzivat pokud mame pouze balik loginu oddelenych strednikem    
                        {
                            string[] spl = identAsString.SplitByChars(";");
                            principals.AddRange(spl.Select(userValue => GetSPPrincipal(web, userValue)));
                        }
                        else
                        {
                            principals.Add(GetSPPrincipal(web, identification));
                        }
                    }
                }
                catch (Exception exc)
                {
                    throw new Exception("GetSPPrincipals exception, users processed (last is the culprit): \n" + processLog.JoinStrings(", ") + "\nOriginal exception:\n" + exc + "\n" + exc.StackTrace);
                }

                principals.RemoveAll(p => p == null); //odstranime vyskyty, ktere se nepodarilo nacist

                return principals.Distinct(new SPPrincipalComparer()) /*.OrderBy(p => p.LoginName)*/.ToList();
                // ReSharper restore PossibleMultipleEnumeration
            }
        }

        public static List<SPUser> GetSPUsersFromADGroup(this SPWeb web, string loginName, List<string> loadedLogins = null, int depth = 0)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (loginName == null) throw new ArgumentNullException("loginName");

            using (new SPMonitoredScope("GetSPUsersFromADGroup: " + loginName))
            {

                bool reachMax;
                List<SPUser> users = new List<SPUser>();
                SPPrincipalInfo[] principals;

                try
                {
                    principals = SPUtility.GetPrincipalsInGroup(web, loginName, 10000, out reachMax);
                }
                catch (InvalidOperationException) // claims doesn't work with GetPrincipalsInGroup
                {
                    return users;
                }

                if (principals == null) return users;

                if (reachMax) throw new Exception("Dosahnut maximalni limit! Pravdepodobne rekurze v AD"); //TODO rekurzi resime zvlast takze to bude jeste neco jineho

                List<string> logins = principals.Select(p => p.LoginName).ToList();

                foreach (string login in logins)
                {
                    SPUser user = web.GetSPUser(login);

                    if (user == null) continue; //uzivatel muze byt disabled pak ho nenactem

                    if (user.IsDomainGroup)
                    {
                        if (loadedLogins == null) loadedLogins = new List<string>();
                        if (loadedLogins.Contains(loginName))
                        {
                            var loadedLoginsInfo = "Recursion usage of Group in Active Directory!!\n";
                            foreach (string loadedLogin in loadedLogins)
                            {
                                SPUser lu = web.GetSPUser(loadedLogin);
                                loadedLoginsInfo += "[" + lu.ID + "] - [ " + lu.Name + "] - [" + lu.LoginName + "]\n";
                            }

                            throw new System.DirectoryServices.ActiveDirectory.ActiveDirectoryObjectExistsException(loadedLoginsInfo);
                        }

                        loadedLogins.Add(loginName);
                        users.AddRange(GetSPUsersFromADGroup(web, login, loadedLogins, depth++));
                        loadedLogins.Remove(loginName);
                    }
                    else
                    {
                        users.Add(user);
                    }
                }

                return users.Distinct(new SPUserComparer()).ToList();
            }
        }

        /// <summary>
        /// Get SPUsers defined by array of identifications. If identification is group, all users from that will be returns
        /// </summary>
        /// <param name="web"></param>
        /// <param name="identification"></param>
        /// <returns></returns>
        public static List<SPUser> GetSPUsers(this SPWeb web, object identification)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (identification == null) throw new ArgumentNullException("identification");

            return GetSPUsers(web, new[] { identification });
        }

        /// <summary>
        /// Get SPUsers defined by array of identifications or straight by string. If identification is group, all users from that will be returns
        /// </summary>
        /// <param name="web"></param>
        /// <param name="identifications"></param>
        /// <returns>List with users or empty list. Never returns null</returns>
        public static List<SPUser> GetSPUsers(this SPWeb web, IEnumerable identifications)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (identifications == null) throw new ArgumentNullException("identifications");

            List<SPUser> users = new List<SPUser>();

            using (new SPMonitoredScope("GetSPUsers"))
            {
                foreach (SPPrincipal principal in GetSPPrincipals(web, identifications))
                {
                    using (new SPMonitoredScope("Process SPPrincipal - " + principal.LoginName))
                    {
                        if (principal is SPUser)
                        {
                            SPUser user = (SPUser) principal;

                            // TODO resit, ze muze byt SP skupina se stejnym jmenem jako AD skupina, a v SP skupine muze byt AD skupina + k ni jeste dalsi uzivatel
                            // a v tu chvili se nacte jen AD skupina, a co k ni je v te SP skupine uz se nezobrazi

                            if (user.IsDomainGroup) //nacteme uzivatele z AD skupiny
                            {
                                users.AddRange(GetSPUsersFromADGroup(web, user.LoginName));
                            }
                            else
                            {
                                users.Add(user); //obycejny SPUSer
                            }
                        }
                        else
                        {
                            SPGroup group = (SPGroup) principal;
                            users.AddRange(group.GetSPUsers()); //SP skupiny rozepiseme
                        }
                    }
                }
            }

            return users.Distinct(new SPUserComparer()).ToList();
        }

        #endregion

        /// <summary>
        /// Checks if we're running in sandbox
        /// </summary>
        /// <param name="web"></param>
        /// <returns></returns>
        public static bool InSandbox(this SPWeb web)
        {
            return AppDomain.CurrentDomain.FriendlyName.Contains("Sandbox");
        }

        /// <summary>
        /// Checks if we're running in OWSTIMER
        /// </summary>
        /// <param name="web"></param>
        /// <returns></returns>
        public static bool InTimerJob(this SPWeb web)
        {
            return System.Diagnostics.Process.GetCurrentProcess().ProcessName.ToLowerInvariant() == "owstimer";
        }

        #region OpenList

        /// <summary>
        /// Tries to open list with given guid - mostly for usage in scripts
        /// </summary>
        /// <param name="web"></param>
        /// <param name="listGuid"></param>
        /// <param name="throwExc"></param>
        /// <returns>list or null</returns>
        /// <exception cref="SPListAccesDeniedException"></exception>
        /// <exception cref="SPListNotFoundException"></exception>
        public static SPList OpenList(this SPWeb web, Guid listGuid, bool throwExc = false)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (listGuid == null) throw new ArgumentNullException("listGuid");
            if (listGuid.IsEmpty()) throw new ArgumentException("Guid is Empty");

            return OpenListFromGuid(web, listGuid, throwExc);
        }

        /// <summary>
        /// Tries to open a list by name, "internal name" (part of url, since there's no official internal name for lists in SP) 
        /// or Guid in string 
        /// </summary>
        /// <param name="web"></param>
        /// <param name="listIdentification"></param>
        /// <param name="throwExc"></param>
        /// <returns>list or null</returns>
        /// <exception cref="SPListAccesDeniedException"></exception>
        /// <exception cref="SPListNotFoundException"></exception>
        public static SPList OpenList(this SPWeb web, string listIdentification, bool throwExc = false)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (listIdentification == null) throw new ArgumentNullException("listIdentification");

            string cacheKey = "OpenlistDict_" + web.Site.ID + web.ID;

            using (new SPMonitoredScope("OpenList - " + listIdentification))
            {
                Guid listGuid;

                #region Open By Guid

                if (Guid.TryParse(listIdentification, out listGuid))
                {
                    return OpenListFromGuid(web, listGuid, throwExc);
                }

                #endregion

                #region Open by InternalName

                string listInternalName = SPListExtensions.ListInternalName(listIdentification.Trim());

                Dictionary<string, Guid> result = (Dictionary<string, Guid>) HttpRuntime.Cache.Get(cacheKey);

                if (result == null)
                {
                    //Nacteme vse pro celej web
                    web.RunElevated(delegate(SPWeb elevatedWeb)
                    {
                        result = elevatedWeb.Lists.Cast<SPList>().ToDictionary(d => d.InternalName(), d => d.ID);
                        HttpRuntime.Cache.Insert(cacheKey, result, null, DateTime.Now.AddDays(1), Cache.NoSlidingExpiration, CacheItemPriority.Normal, null);
                        return null;
                    });
                }

                if (result.ContainsKey(listInternalName))
                {
                    return OpenListFromGuid(web, result[listInternalName], throwExc);
                }

                #endregion

                if (!throwExc) return null;
                throw new SPListNotFoundException(listIdentification, web);
            }
        }

        public static SPList OpenListFromGuid(SPWeb web, Guid listGuid, bool throwExc)
        {
            try
            {
                bool originalCatchValue = SPSecurity.CatchAccessDeniedException;
                SPSecurity.CatchAccessDeniedException = false;

                SPList list = web.Lists[listGuid];

                SPSecurity.CatchAccessDeniedException = originalCatchValue;
                return list;
            }
            catch (UnauthorizedAccessException)
            {
                if (throwExc)
                {
                    throw new SPListAccesDeniedException(listGuid, web);
                }
            }
            catch (Exception) {}

            if (throwExc)
            {
                throw new SPListNotFoundException(listGuid, web);
            }

            return null;
        }

        #endregion

        public static void RunWithAllowUnsafeUpdates(this SPWeb web, Action func)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (func == null) throw new ArgumentNullException("func");

            if (web.AllowUnsafeUpdates) func();
            else
            {
                web.AllowUnsafeUpdates = true;
                func();
                web.AllowUnsafeUpdates = false;
            }
        }

        /// <summary>
        /// Runs code as elevated admin
        /// </summary>
        /// <param name="web"></param>
        /// <param name="func"></param>
        /// <returns></returns>
        public static object RunElevated(this SPWeb web, Func<SPWeb, object> func)
        {
            if (web == null) throw new ArgumentNullException("web");
            if (func == null) throw new ArgumentNullException("func");

            Guid siteGuid = web.Site.ID;
            Guid webGuid = web.ID;
            object result = null;

            if (web.InSandbox() || web.CurrentUser.IsSiteAdmin) //Don't runeleveted if user is admin or already run with admin rights
            {
                bool origSiteunsafe = web.Site.AllowUnsafeUpdates;
                web.Site.AllowUnsafeUpdates = true;
                web.RunWithAllowUnsafeUpdates(delegate { result = func(web); });
                web.Site.AllowUnsafeUpdates = origSiteunsafe;
            }
            else
            {
                using (new SPMonitoredScope("RunElevated SPWeb - " + web.Url))
                {
                    SPSecurity.RunWithElevatedPrivileges(delegate
                    {
                        using (SPSite elevatedSite = new SPSite(siteGuid, web.Site.GetSystemToken()))
                        {
                            elevatedSite.AllowUnsafeUpdates = true;
                            using (SPWeb elevatedWeb = elevatedSite.OpenW(webGuid, true))
                            {
                                elevatedWeb.AllowUnsafeUpdates = true;
                                result = func(elevatedWeb);
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