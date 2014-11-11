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