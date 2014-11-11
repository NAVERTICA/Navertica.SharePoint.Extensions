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
using System.Collections;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Web;
using System.Web.Caching;
using Microsoft.SharePoint;

namespace Navertica.SharePoint.Extensions
{
    public static class SPPrincipalExtensions
    {
        /// <summary>
        /// Checks, if the current user is enabled in Active Directory
        /// </summary>
        /// <param name="user"></param>
        /// <param name="de"></param>
        /// <returns></returns>
        public static bool? Enabled(this SPUser user, DirectoryEntry de)
        {
            if (user == null) throw new ArgumentNullException("user");

            //http://support.microsoft.com/kb/305144
            /*
                512 - Enable Account
                514 - Disable account
                544 - Account Enabled - Require user to change password at first logon
                4096 - Workstation/server
                66048 - Enabled, password never expires
                66050 - Disabled, password never expires
                262656 - Smart Card Logon Required
                532480 - Domain controller
             */
            object prop = user.GetUserProperty(de, "useraccountcontrol");

            try
            {
                int val = Convert.ToInt32(prop);

                if (val == 512 || val == 66048) return true;
                if (val == 514 || val == 66050) return false;
            }
            catch
            {
                return null;
            }

            return null;
        }

        /// <summary>
        /// Returns guids of all the webs this user can access
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public static IEnumerable<Guid> FindWebsForUser(this SPUser user)
        {
            if (user == null) throw new ArgumentNullException("user");

            List<Guid> guids = new List<Guid>();

            foreach (SPWeb w in user.ParentWeb.Site.RootWeb.GetSubwebsForCurrentUser())
            {
                guids.Add(w.ID);
                w.Dispose();
            }

            return guids;

        }

        /// <summary>
        /// Returns integer IDs of all the principals (users/groups)
        /// </summary>
        /// <param name="principals"></param>
        /// <returns></returns>
        public static List<int> GetIds(this IEnumerable<SPPrincipal> principals)
        {
            if (principals == null) throw new ArgumentNullException("principals");

            return principals.Select(principal => principal.ID).ToList();
        }


        /// <summary>
        /// LoginName in lowercase, without Claims prefix
        /// </summary>
        /// <param name="principal"></param>
        /// <returns></returns>
        public static string LoginNameNormalized(this SPPrincipal principal)
        {
            if (principal == null) throw new ArgumentNullException("principal");

            return ( principal.LoginName.Contains('|') ? principal.LoginName.Split('|')[1] : principal.LoginName ).ToLowerInvariant();
        }

        /// <summary>
        /// Returns LoginNames in lowercase, without Claims prefix, for all the principals (users/groups)
        /// </summary>
        /// <param name="principals"></param>
        /// <returns></returns>
        public static List<string> GetLogins(this IEnumerable<SPPrincipal> principals)
        {
            if (principals == null) throw new ArgumentNullException("principals");

            return principals.Select(principal => principal.LoginNameNormalized()).ToList();
        }


        /// <summary>
        /// Tries to load this user's manager from ActiveDirectory
        /// </summary>
        /// <param name="user"></param>
        /// <param name="de"></param>
        /// <returns></returns>
        public static SPUser GetManager(this SPUser user, DirectoryEntry de)
        {
            if (user == null) throw new ArgumentNullException("user");
            if (de == null) throw new ArgumentNullException("de");
            if (!IsValid(de)) throw new DirectoryEntryException(de);

            object value = GetUserProperty(user, de, "Manager");

            if (value == null) return null;

            string[] props = value.ToString().Split(',');
            string managerName = props.FirstOrDefault(s => s.ToLowerInvariant().StartsWith("cn"));

            if (managerName == null) return null;

            managerName = managerName.Split('=')[1];
            SPUser manager = user.ParentWeb.GetSPUser(managerName);

            return manager;
        }

        /// <summary>
        /// Returns the names for all the principals (users/groups)
        /// </summary>
        /// <param name="principals"></param>
        /// <returns></returns>
        public static List<string> GetNames(this IEnumerable<SPPrincipal> principals)
        {
            if (principals == null) throw new ArgumentNullException("principals");

            return principals.Select(principal => principal.Name).ToList();
        }


        /// <summary>
        /// Returns the SPFieldUserValue to be saved in a SPFieldUser for the current principal (user/group)
        /// </summary>
        /// <param name="principal"></param>
        /// <returns></returns>
        public static SPFieldUserValue GetSPFieldUserValue(this SPPrincipal principal)
        {
            if (principal == null) throw new ArgumentNullException("principal");

            return new SPFieldUserValue(principal.ParentWeb, GetSPFieldUserValueFormat(principal));
        }      

        /// <summary>
        /// Returns the text format "ID;#Name" for the current principal (user/group)
        /// </summary>
        /// <param name="principal"></param>
        /// <returns></returns>
        public static string GetSPFieldUserValueFormat(this SPPrincipal principal)
        {
            if (principal == null) throw new ArgumentNullException("principal");

            return principal.ID + ";#" + principal.Name;
        }

        /// <summary>
        /// Returns the SPFieldUserValueCollection to be saved in a SPFieldUser for the principals (users/groups)
        /// </summary>
        /// <param name="principals"></param>
        /// <returns></returns>
        public static SPFieldUserValueCollection GetSPFieldUserValueCollection(this IEnumerable<SPPrincipal> principals)
        {
            if (principals == null) throw new ArgumentNullException("principals");

            SPFieldUserValueCollection result = new SPFieldUserValueCollection();

            foreach (SPPrincipal princ in principals.Where(p => p != null))
            {
                result.Add(new SPFieldUserValue(princ.ParentWeb, princ.ID, princ.Name));
            }

            return result;
        }

        /// <summary>
        /// Get all users from an SPGroup, expands Active Directory groups.
        /// </summary>
        /// <param name="group"></param>
        /// <returns>List of SPUsers or empty List. Never returns null.</returns>
        public static List<SPUser> GetSPUsers(this SPGroup group)
        {
            if (group == null) throw new ArgumentNullException("group");

            string cacheKey = string.Format("GetSPUsersFromSPGroup_{0}", group.LoginName);
            List<SPUser> spUsers = (List<SPUser>) HttpRuntime.Cache.Get(cacheKey);

            if (spUsers != null && !group.ParentWeb.InTimerJob())
            {
                return new List<SPUser>(spUsers);
            }

            if (!group.CanCurrentUserViewMembership) // if not possible, use elevated rights
            {
                List<int> ids = new List<int>();
                group.ParentWeb.RunElevated(delegate(SPWeb elevWeb)
                {
                    SPGroup g = elevWeb.GetSPGroup(group.ID);
                    ids = GetUsersFromCollection(g.Users).Select(u => u.ID).ToList();
                    return null;
                });
                spUsers = group.ParentWeb.GetSPUsers(ids).ToList();
            }
            else
            {
                spUsers = GetUsersFromCollection(group.Users);
            }

            spUsers.RemoveAll(user => user.ID == group.ParentWeb.Site.SystemAccount.ID); // siteadmin is in every group by default, we don't want this

            HttpRuntime.Cache.Insert(cacheKey, spUsers, null, DateTime.Now.AddMinutes(30), Cache.NoSlidingExpiration);

            return new List<SPUser>(spUsers);
        }

        /// <summary>
        /// Returns all the users in given collection - expands groups, including Active Directory groups.
        /// </summary>
        /// <param name="userCollection"></param>
        /// <returns></returns>
        private static List<SPUser> GetUsersFromCollection(SPUserCollection userCollection)
        {
            List<SPUser> spUsers = new List<SPUser>();
            foreach (SPUser user in userCollection)
            {
                if (user.IsDomainGroup) // domain group inside a SharePoint group
                {
                    spUsers.AddRange(user.ParentWeb.GetSPUsersFromADGroup(user.LoginName));
                }
                else
                {
                    spUsers.Add(user);
                }
            }

            return spUsers.Distinct(new SPUserComparer()).ToList();
        }

        /// <summary>
        /// Load the property of given name for given user from Active Directory
        /// </summary>
        /// <param name="user"></param>
        /// <param name="de"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public static object GetUserProperty(this SPUser user, DirectoryEntry de, string propertyName)
        {
            if (user == null) throw new ArgumentNullException("user");
            if (de == null) throw new ArgumentNullException("de");
            if (propertyName == null) throw new ArgumentNullException("propertyName");
            if (!IsValid(de)) throw new DirectoryEntryException(de);

            propertyName = propertyName.ToLowerInvariant();
            DictionaryNVR results = GetUserProperties(user, de, new[] { propertyName });
            if (results != null && results.ContainsKey(propertyName)) return results[propertyName];

            return null;
        }

        /// <summary>
        /// Loads the properties from Active Directory into dictionary. Keys in dictionary will be lowercase.
        /// </summary>
        /// <param name="user"></param>
        /// <param name="de"></param>
        /// <param name="properties"></param>
        /// <returns></returns>
        public static DictionaryNVR GetUserProperties(this SPUser user, DirectoryEntry de, IEnumerable<string> properties = null)
        {
            if (user == null) throw new ArgumentNullException("user");
            if (de == null) throw new ArgumentNullException("de");
            if (!IsValid(de)) throw new DirectoryEntryException(de);

            properties = properties == null ? new List<string>() : properties.ToLowerInvariant().ToList();

            string loginName = user.LoginNameNormalized();
            if (loginName.Contains('\\'))
            {
                loginName = loginName.Split('\\').Last();
            }

            SearchResultCollection results = LDAP_Find(de, "(&(&(objectCategory=user)(|(cn=" + user.Name + ")(sAMAccountName=" + loginName + "))))", properties);

            if (results.Count != 1)
            {
                return null;
            }

            DictionaryNVR dict = new DictionaryNVR();

            SearchResult userProperties = results[0];
            List<string> keys = userProperties.Properties.Cast<DictionaryEntry>().Select(dictEntry => dictEntry.Key.ToString()).ToList();

            List<string> notFound = properties.Except(keys).ToList();

            foreach (string key in keys)
            {
                if (userProperties.Properties[key].Count == 1)
                {
                    dict.Add(key, userProperties.Properties[key][0]);
                }
                else
                {
                    List<object> l = new List<object>();
                    foreach (object o in userProperties.Properties[key])
                    {
                        l.Add(o);
                    }
                    dict.Add(key, l);
                }
            }

            foreach (string not in notFound)
            {
                dict.Add(not, null);
            }

            dict.Sort();
            return dict;
        }
        
        private static SearchResultCollection LDAP_Find(DirectoryEntry de, string filter, IEnumerable<string> properties)
        {
            string ldapString = "-";
            try
            {
                DirectorySearcher ds = new DirectorySearcher(de);
                ds.Filter = filter;

                ldapString = de.Name + " - " + ds.Filter;

                ds.PropertiesToLoad.AddRange(properties.ToArray());

                return ds.FindAll();
            }
            catch (Exception ex)
            {
                throw new Exception("LDAP_Find problem\n" + ldapString + "\n" + ex + "\n");
            }
        }    

        public static bool IsValid(this DirectoryEntry de)
        {
            if (de == null) return false;

            try
            {
                // ReSharper disable UnusedVariable
                Guid guid = de.Guid; //cause an exception
                // ReSharper restore UnusedVariable
                return true;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        ///  Get this user's preferred language LCID
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public static int GetPreferredLanguage(this SPUser user)
        {
            var preflangs = user.LanguageSettings.PreferredDisplayLanguages;
            int lang = preflangs.Count > 0 ? preflangs[0].LCID : (int)user.ParentWeb.Language;
            if (lang < 1000) lang = (int)user.ParentWeb.Language;
            return lang;
        }
    }
}