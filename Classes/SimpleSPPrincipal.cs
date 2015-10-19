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

using Microsoft.SharePoint;
using Navertica.SharePoint.Extensions;

namespace Navertica.SharePoint
{
    /// <summary>
    /// Simple class to hold SharePoint principal info 
    /// </summary>
    public class SimpleSPPrincipal
    {
        private int _id;
        private string _loginName;
        private string _name;
        private string _email;

        public int Id { get { return _id; } }
        public string LoginName { get { return _loginName; } }
        public string Name { get { return _name; } }
        public string Email { get { return _email; } }

        public SimpleSPPrincipal(SPPrincipal principal)
        {
            _id = principal.ID;
            _loginName = principal.LoginNameNormalized();
            _name = principal.Name;
            _email = principal is SPUser ? ((SPUser)principal).Email : null;
        }

        public SPPrincipal GetSPPrincipal(SPWeb web)
        {
            return web.GetSPPrincipal(_loginName);
        }
    }
}