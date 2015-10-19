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
using System.Collections.Generic;
using Microsoft.SharePoint;

namespace Navertica.SharePoint
{
    /// <summary>
    /// Simple class to hold SharePoint Lookup info
    /// </summary>
    public class SimpleSPLookup
    {
        private int _id;
        private string _value;
        private WebListItemId _wli;
        public SimpleSPLookup(SPFieldLookup field, SPFieldLookupValue val)
        {
            _id = val.LookupId;
            _value = val.LookupValue;
            _wli = new WebListItemId(field.ParentList.ParentWeb.ID, new Guid(field.LookupList), _id);
        }

        public int Id { get { return _id; } }
        public string Value { get { return _value; } }
        public WebListItemId WLI { get { return _wli; } }

        public SPFieldLookupValue GetSPFieldLookupValue()
        {
            return new SPFieldLookupValue(_id, _value);
        }
    }    
}

namespace Navertica.SharePoint.Extensions
{
    public static class SimpleSPLookupListExtensions
    {
        public static SPFieldLookupValueCollection GetSPFieldLookupValues(this List<SimpleSPLookup> lookups)
        {
            SPFieldLookupValueCollection coll = new SPFieldLookupValueCollection();
            foreach (var lookup in lookups)
            {
                coll.Add(lookup.GetSPFieldLookupValue());
            }
            return coll;
        }
    }
}