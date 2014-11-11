using System.Collections.Generic;
using Microsoft.SharePoint;

// ReSharper disable once CheckNamespace
namespace Navertica.SharePoint
{
    public class SPPrincipalComparer : IEqualityComparer<SPPrincipal>
    {
        public bool Equals(SPPrincipal x, SPPrincipal y)
        {
            try
            {
                return x.LoginName == y.LoginName;
            }
            catch
            {
                return false;
            }
        }

        public int GetHashCode(SPPrincipal obj)
        {
            return obj.ID;
        }
    }

    public class SPUserComparer : IEqualityComparer<SPUser>
    {
        public bool Equals(SPUser x, SPUser y)
        {
            try
            {
                return x.LoginName == y.LoginName;
            }
            catch
            {
                return false;
            }
        }

        public int GetHashCode(SPUser obj)
        {
            return obj.ID;
        }
    }
}