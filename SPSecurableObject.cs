using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint;

namespace Navertica.SharePoint.Extensions
{
    public static class SPSecurableObjectExtensions
    {
        public static bool CopyRights(this SPSecurableObject securableObject, SPSecurableObject toSecurableObject, bool deleteOldRights)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");
            if (toSecurableObject == null) throw new ArgumentNullException("toSecurableObject");

            try
            {
                securableObject.RunElevated(delegate(SPSecurableObject securableObjectElevated)
                {
                    toSecurableObject.RunElevated(delegate(SPSecurableObject toSecurableObjectElevated)
                    {
                        if (!toSecurableObjectElevated.HasUniqueRoleAssignments) toSecurableObjectElevated.BreakRoleInheritance(true);

                        if (deleteOldRights)
                        {
                            toSecurableObjectElevated.RemoveRights();
                        }

                        foreach (SPRoleAssignment roleAssignment in securableObjectElevated.RoleAssignments)
                        {
                            toSecurableObjectElevated.RoleAssignments.Add(roleAssignment);
                        }

                        return null;
                    });
                    return null;
                });

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Gets information about securableObject
        /// </summary>
        /// <param name="securableObject"></param>
        /// <returns></returns>
        public static string GetInfo(this SPSecurableObject securableObject)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");

            switch (securableObject.GetType().Name)
            {
                case "SPListItem":
                    return "Securable object : SPListItem - " + ( (SPListItem) securableObject ).FormUrlDisplay();
                case "SPList":
                    return "Securable object : SPList - " + ( (SPList) securableObject ).DefaultViewUrl;
                case "SPWeb":
                    return "Securable object : SPWeb - " + ( (SPWeb) securableObject ).Url;
                default:
                    return "";
            }
        }

        /// <summary>
        /// Get permissions of securableObject in JSON format
        /// </summary>
        /// <param name="securableObject"></param>
        /// <returns></returns>
        public static string GetPermissionsJSON(this SPSecurableObject securableObject)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");

            List<string> results = new List<string>();

            securableObject.RunElevated(delegate(SPSecurableObject obj)
            {
                try
                {
                    foreach (SPRoleAssignment role in obj.RoleAssignments)
                    {
                        List<string> roledef = role.RoleDefinitionBindings.Cast<SPRoleDefinition>().Select(rdef => rdef.Name).ToList();
                        bool group = role.Member is SPGroup;
                        results.Add("{ \"" + ( group ? "group" : "user" ) + "\": \"" + role.Member.LoginName + "\", \"rights\": \"" + roledef.JoinStrings(", ") + "\" }");
                    }
                }
                catch
                {
                    try
                    {
                        results.Add("{ \"user\": \"" + obj.GetWeb().CurrentUser.LoginName + "\", \"rights\": \"none\" }");
                    }
                    catch
                    {
                        results.Add(( "GetPermissionsJSON problem - no SPContext" ));
                    }
                }

                return null;
            });

            return "[" + results.JoinStrings(", ") + "]";
        }

        public static SPRoleDefinition GetRoleDefinition(this SPSecurableObject securableObject, SPRoleType roleType)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");

            switch (securableObject.GetType().Name)
            {
                case "SPListItem":
                    return ( (SPListItem) securableObject ).Web.RoleDefinitions.GetByType(roleType);
                case "SPList":
                    return ( (SPList) securableObject ).ParentWeb.RoleDefinitions.GetByType(roleType);
                case "SPWeb":
                    return ( (SPWeb) securableObject ).RoleDefinitions.GetByType(roleType);
            }
            return null;
        }

        /// <summary>
        /// Gets o roledefinition by name
        /// </summary>
        /// <param name="securableObject"></param>
        /// <param name="roleDefinitonName"></param>
        /// <returns>roledefinition or null</returns>
        public static SPRoleDefinition GetRoleDefinition(this SPSecurableObject securableObject, string roleDefinitonName)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");
            if (roleDefinitonName == null) throw new ArgumentNullException("roleDefinitonName");

            return securableObject.GetWeb().RoleDefinitions.Cast<SPRoleDefinition>().SingleOrDefault(def => def.Name == roleDefinitonName.Trim());
        }

        /// <summary>
        /// Gets the admin's email
        /// </summary>
        /// <param name="securableObject"> </param>
        /// <returns></returns>
        public static string GetSystemAccountEmail(this SPSecurableObject securableObject)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");

            switch (securableObject.GetType().Name)
            {
                case "SPListItem":
                    return ( (SPListItem) securableObject ).ParentList.ParentWeb.Site.SystemAccount.Email;
                case "SPList":
                    return ( (SPList) securableObject ).ParentWeb.Site.SystemAccount.Email;
                case "SPWeb":
                    return ( (SPWeb) securableObject ).Site.SystemAccount.Email;
                default:
                    return "";
            }
        }

        public static SPWeb GetWeb(this SPSecurableObject securableObject)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");

            switch (securableObject.GetType().Name)
            {
                case "SPListItem":
                    return ( (SPListItem) securableObject ).Web;
                case "SPList":
                    return ( (SPList) securableObject ).ParentWeb;
                case "SPWeb":
                    return ( (SPWeb) securableObject );
            }
            return null;
        }

        /// <summary>
        /// Run code as elevated admin
        /// </summary>
        /// <param name="securableObject"></param>
        /// <param name="func"></param>
        /// <returns></returns>
        public static object RunElevated(this SPSecurableObject securableObject, Func<SPSecurableObject, object> func)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");
            if (func == null) throw new ArgumentNullException("func");

            object result = null;

            switch (securableObject.GetType().Name)
            {
                case "SPListItem":
                    ( (SPListItem) securableObject ).RunElevated(delegate(SPListItem elevatedItem)
                    {
                        result = func(elevatedItem);
                        return null;
                    });
                    break;
                case "SPList":
                    ( (SPList) securableObject ).RunElevated(delegate(SPList elevatedList)
                    {
                        result = func(elevatedList);
                        return null;
                    });
                    break;
                case "SPWeb":
                    ( (SPWeb) securableObject ).RunElevated(delegate(SPWeb elevatedWeb)
                    {
                        result = func(elevatedWeb);
                        return null;
                    });
                    break;
            }
            return result;
        }

        #region Remove

        /// <summary>
        /// Smazat aktuální role assignments
        /// </summary>
        /// <param name="securableObject"></param>
        public static bool RemoveRights(this SPSecurableObject securableObject)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");

            try
            {
                securableObject.RunElevated(delegate(SPSecurableObject securableObjectElevated)
                {
                    securableObjectElevated.BreakRoleInheritance(true);

                    for (int i = securableObjectElevated.RoleAssignments.Count - 1; i > -1; i--) // smazat aktuální role assignments
                    {
                        securableObjectElevated.RoleAssignments.Remove(i);
                    }
                    return null;
                });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Smaze roli pro dany objekt. Nefunguje na SPRoleType.Guest
        /// </summary>
        /// <param name="securableObject"> </param>
        /// <param name="roleDefinition"></param>
        /// <returns></returns>
        public static bool RemoveRights(this SPSecurableObject securableObject, SPRoleDefinition roleDefinition)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");
            if (roleDefinition == null) throw new ArgumentNullException("roleDefinition");

            try
            {
                securableObject.RunElevated(delegate(SPSecurableObject securableObjectElevated)
                {
                    securableObjectElevated.BreakRoleInheritance(true);

                    for (int i = securableObjectElevated.RoleAssignments.Count - 1; i > -1; i--) // smazat aktuální role assignments
                    {
                        SPRoleAssignment assignment = securableObjectElevated.RoleAssignments[i];
                        SPRoleDefinitionBindingCollection col = assignment.RoleDefinitionBindings;

                        foreach (SPRoleDefinition currentRoleDefinition in col)
                        {
                            if (roleDefinition.Id == currentRoleDefinition.Id)
                            {
                                col.Remove(roleDefinition);
                            }
                        }
                        assignment.Update();
                    }

                    return null;
                });

                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool RemoveRights(this SPSecurableObject securableObject, SPRoleType roleType)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");

            SPRoleDefinition roleDefinition = securableObject.GetRoleDefinition(roleType);
            return RemoveRights(securableObject, roleDefinition);
        }

        public static bool RemoveRights(this SPSecurableObject securableObject, SPPrincipal principal)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");
            if (principal == null) throw new ArgumentNullException("principal");

            return RemoveRights(securableObject, new[] { principal });
        }

        public static bool RemoveRights(this SPSecurableObject securableObject, IEnumerable<SPPrincipal> principals)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");
            if (principals == null) throw new ArgumentNullException("principals");

            try
            {
                securableObject.RunElevated(delegate(SPSecurableObject securableObjectElevated)
                {
                    securableObjectElevated.BreakRoleInheritance(true);

                    principals = principals.ToList();

                    for (int i = securableObjectElevated.RoleAssignments.Count - 1; i > -1; i--) // smazat aktuální role assignments
                    {
                        foreach (SPPrincipal principal in principals)
                        {
                            if (securableObjectElevated.RoleAssignments[i].Member.ID == principal.ID)
                            {
                                securableObjectElevated.RoleAssignments.Remove(i);
                            }
                        }
                    }
                    return null;
                });

                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Set

        public static bool SetRights(this SPSecurableObject securableObject, SPPrincipal principal, SPRoleDefinition roleDefinition, bool deleteCurrent = false)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");
            if (principal == null) throw new ArgumentNullException("principal");
            if (roleDefinition == null) throw new ArgumentNullException("roleDefinition");

            return SetRights(securableObject, new[] { principal }, roleDefinition, deleteCurrent);
        }

        public static bool SetRights(this SPSecurableObject securableObject, SPPrincipal principal, SPRoleType roleType, bool deleteCurrent = false)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");
            if (principal == null) throw new ArgumentNullException("principal");

            SPRoleDefinition roleDefinition = securableObject.GetRoleDefinition(roleType);
            return SetRights(securableObject, new[] { principal }, roleDefinition, deleteCurrent);
        }

        public static bool SetRights(this SPSecurableObject securableObject, IEnumerable<SPPrincipal> principals, SPRoleType roleType, bool deleteCurrent = false)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");
            if (principals == null) throw new ArgumentNullException("principals");

            SPRoleDefinition roleDefinition = securableObject.GetRoleDefinition(roleType);
            return SetRights(securableObject, principals, roleDefinition, deleteCurrent);
        }

        /// <summary>
        /// Nastaví uživateli individuální přístupová práva pro položku, bude fungovat pod jakýmkoliv uživatelem - tzn. pozor
        /// </summary>
        /// <param name="securableObject">položka, ke které má uživatel získat práva</param>
        /// <param name="principals">uživatel/sp skupina</param>
        /// <param name="roleDefinition"> práva, která má uživatel získat</param>
        /// <param name="deleteCurrent"> </param>
        /// <returns>Vrací false, když něco selže, chyba bude v logu</returns>
        public static bool SetRights(this SPSecurableObject securableObject, IEnumerable<SPPrincipal> principals, SPRoleDefinition roleDefinition, bool deleteCurrent = false)
        {
            if (securableObject == null) throw new ArgumentNullException("securableObject");
            if (principals == null) throw new ArgumentNullException("principals");
            if (roleDefinition == null) throw new ArgumentNullException("roleDefinition");
            
            var spPrincipals = principals as IList<SPPrincipal> ?? principals.ToList();
            if (!spPrincipals.Any()) return false;

            securableObject.RunElevated(delegate(SPSecurableObject securableObjectElevated)
            {
                if (deleteCurrent) securableObject.RemoveRights();

                if (!securableObjectElevated.HasUniqueRoleAssignments) securableObjectElevated.BreakRoleInheritance(true);

                foreach (SPRoleAssignment roleAssignment in spPrincipals.Select(principal => new SPRoleAssignment(principal)))
                {
                    roleAssignment.RoleDefinitionBindings.Add(roleDefinition);
                    securableObjectElevated.RoleAssignments.Add(roleAssignment);
                }

                return null;
            });

            return true;
        }

        #endregion
    }
}