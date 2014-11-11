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
        /// Returns guids of webs which user can access
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

            /*
             //puvodni kod asi relikt 2007
            try
            {
                string userLogin = user.LoginName;

                user.ParentWeb.Site.RunElevated(delegate(SPSite elevatedSite)
                {
                    foreach (SPWeb web in elevatedSite.AllWebs)
                    {
                        //SPUser user = web.GetSPUser(userLogin);
                        if (user.IsSiteAdmin || web.DoesUserHavePermissions(user.LoginName, SPBasePermissions.Open))
                        {
                            guids.Add(web.ID);
                        }
                        web.Dispose();
                    }
                    return null;
                });

                return guids;
            }
            catch (Exception exc)
            {
                Tools.Log("FindWebsForUser\n" + exc);
                return null;
            }
             * */
        }

        public static List<int> GetIds(this IEnumerable<SPPrincipal> principals)
        {
            if (principals == null) throw new ArgumentNullException("principals");

            return principals.Select(principal => principal.ID).ToList();
        }

        public static List<int> GetIds(this IEnumerable<SPUser> users)
        {
            if (users == null) throw new ArgumentNullException("users");

            return users.Select(user => user.ID).ToList();
        }

        /// <summary>
        /// LoginName without Claims
        /// </summary>
        /// <param name="principal"></param>
        /// <returns></returns>
        public static string LoginNameNormalized(this SPPrincipal principal)
        {
            if (principal == null) throw new ArgumentNullException("principal");

            return ( principal.LoginName.Contains('|') ? principal.LoginName.Split('|')[1] : principal.LoginName ).ToLowerInvariant();
        }

        public static List<string> GetLogins(this IEnumerable<SPPrincipal> principals)
        {
            if (principals == null) throw new ArgumentNullException("principals");

            return principals.Select(principal => principal.LoginNameNormalized()).ToList();
        }

        public static List<string> GetLogins(this IEnumerable<SPUser> users)
        {
            if (users == null) throw new ArgumentNullException("users");

            return users.Cast<SPPrincipal>().GetLogins();
        }

        public static SPUser GetManager(this SPUser user, DirectoryEntry de)
        {
            if (user == null) throw new ArgumentNullException("user");
            if (de == null) throw new ArgumentNullException("de");
            if (!IsValid(de)) throw new DirectoryEntryException(de);

            object value = GetUserProperty(user, de, "Manager");

            if (value == null) return null;

            string[] props = value.ToString().Split(',');
            string managerName = props.Where(s => s.ToLowerInvariant().StartsWith("cn")).FirstOrDefault();

            if (managerName == null) return null;

            managerName = managerName.Split('=')[1];
            SPUser manager = user.ParentWeb.GetSPUser(managerName);

            return manager;
        }

        public static List<string> GetNames(this IEnumerable<SPPrincipal> principals)
        {
            if (principals == null) throw new ArgumentNullException("principals");

            return principals.Select(principal => principal.Name).ToList();
        }

        public static List<string> GetNames(this IEnumerable<SPUser> users)
        {
            if (users == null) throw new ArgumentNullException("users");

            return users.Select(user => user.Name).ToList();
        }

        public static SPFieldUserValue GetSPFieldUserValue(this SPPrincipal principal)
        {
            if (principal == null) throw new ArgumentNullException("principal");

            return new SPFieldUserValue(principal.ParentWeb, GetSPFieldUserValueFormat(principal));
        }

        public static SPFieldUserValue GetSPFieldUserValue(this SPUser user)
        {
            if (user == null) throw new ArgumentNullException("user");

            return GetSPFieldUserValue((SPPrincipal) user);
        }

        public static string GetSPFieldUserValueFormat(this SPPrincipal principal)
        {
            if (principal == null) throw new ArgumentNullException("principal");

            return principal.ID + ";#" + principal.Name;
        }

        public static string GetSPFieldUserValueFormat(this SPUser user)
        {
            if (user == null) throw new ArgumentNullException("user");

            return user.ID + ";#" + user.Name;
        }

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

        public static SPFieldUserValueCollection GetSPFieldUserValueCollection(this IEnumerable<SPUser> users)
        {
            if (users == null) throw new ArgumentNullException("users");

            return GetSPFieldUserValueCollection(users.Cast<SPPrincipal>());
        }

        /// <summary>
        /// Get all users from SPGroup
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

            if (!group.CanCurrentUserViewMembership) //pokud nemuze musime pod zvysenymi pravy
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

            spUsers.RemoveAll(user => user.ID == group.ParentWeb.Site.SystemAccount.ID); // siteadmin je v každé skupině, ale to my nechceme

            HttpRuntime.Cache.Insert(cacheKey, spUsers, null, DateTime.Now.AddMinutes(30), Cache.NoSlidingExpiration);

            return new List<SPUser>(spUsers);
        }

        private static List<SPUser> GetUsersFromCollection(SPUserCollection userCollection)
        {
            List<SPUser> spUsers = new List<SPUser>();
            foreach (SPUser user in userCollection)
            {
                if (user.IsDomainGroup) // domenova skupina v sharepointove skupine
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
        /// Get the property for user from AD
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
        /// Gets properties for user from AD. Keys are in LowerCase format
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

        #region LDAP Functions

        /*
        private static string GetLdapFilter(string name, bool isDomainGroup)
        {
            if (isDomainGroup)
            {
                if (name.Contains('\\'))
                {
                    name = name.Split('\\')[1];
                }
                return "(&(objectCategory=group)(|(cn=" + name + ")(sAMAccountName=" + name + ")))";
            }
            else
            {
                //Pro nacitani props uz nemusime filtrovat pouze enabled users - usery nacitame pomoci SPUtility
                return "(&(&(objectCategory=user)(|(cn=" + name + ")(sAMAccountName=" + name + "))))";
                //return "(&(&(objectCategory=user)(|(cn=" + name + ")(sAMAccountName=" + name + ")))(!(userAccountControl:1.2.840.113556.1.4.803:=2)))";
            }
        }
         * */

        /*
        private static List<SearchResultCollection> LDAP_FindUsersForGroup(this SPUser adGroup, DirectoryEntry de)
        {
            if (adGroup == null) throw new ArgumentNullException("adGroup");
            if (de == null) throw new ArgumentNullException("de");
            if (!IsValid(de)) throw new DirectoryEntryException(de);

            string groupADName = adGroup.LoginName.Split('\\')[1];
            return LDAP_FindUsersForGroup(de, groupADName, new List<string>());
        }
        */

        /*
        private static List<SearchResultCollection> LDAP_FindUsersForGroup(DirectoryEntry de, string groupname, ICollection<string> recursionProtection)
        {
            string[] propertiesToLoad = { "sAMAccountName", "distinguishedName" };

            SearchResultCollection group = LDAP_Find(de, GetLdapFilter(groupname, true), propertiesToLoad);
            List<SearchResultCollection> results = new List<SearchResultCollection>();

            if (group.Count == 0)
            {
                //Tools.Log("LDAP_FindUsersForGroup - skupina " + groupname + " nenalezena\n\n" + Tools.CurrentStackTrace());
                return null;
            }
            if (group.Count > 1)
            {
                //Tools.Log("LDAP_FindUsersForGroup - skupina " + groupname + " nalezena vic nez jednou\n\n" + Tools.CurrentStackTrace());
                return null;
            }

            string groupDistingName = group[0].Properties["distinguishedName"][0].ToString();

            // najit vnorene skupiny a rekurzivne se do nich pustit
            // TODO hlidat nekonecnou rekurzi, jestli se to v AD muze stat

            foreach (SearchResult grp in LDAP_Find(de, "(&(objectCategory=group)(memberOf=" + groupDistingName + "))", propertiesToLoad))
            {
                string insideGroupName = grp.Properties["distinguishedName"][0].ToString();
                string accName = grp.Properties["sAMAccountName"][0].ToString();

                if (recursionProtection.Contains(insideGroupName)) continue;

                recursionProtection.Add(insideGroupName);

                List<SearchResultCollection> resultUsers = LDAP_FindUsersForGroup(de, accName, recursionProtection);
                if (resultUsers != null)
                {
                    results.AddRange(resultUsers);
                }
            }
            // cast s userAccountControl vybira ucty, ktere nejsou disabled
            results.Add(LDAP_Find(de, "(&(&(objectCategory=user)(memberOf=" + groupDistingName + "))(!(userAccountControl:1.2.840.113556.1.4.803:=2)))", propertiesToLoad));

            return results;
        }
        */

        /*
        private static SearchResultCollection LDAP_FindUsers(DirectoryEntry de, string filter)
        {
            return de.LDAP_Find(filter, new[] { "sAMAccountName", "distinguishedName", "cn", "displayName", "Manager", "userAccountControl", "department" });
        }
        */

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

        #endregion

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

        public static int GetPreferredLanguage(this SPUser user)
        {
            var preflangs = user.LanguageSettings.PreferredDisplayLanguages;
            int lang = preflangs.Count > 0 ? preflangs[0].LCID : (int)user.ParentWeb.Language;
            if (lang < 1000) lang = (int)user.ParentWeb.Language;
            return lang;
        }
    }
}