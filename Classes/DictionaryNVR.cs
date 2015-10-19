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
using System.Linq;

namespace Navertica.SharePoint
{
    /// <summary>
    /// simple Dictionary[string, object] that reports the names of nonexistent keys in thrown KeyNotFoundException
    /// </summary>
    public class DictionaryNVR : Dictionary<string, object>
    {
        public DictionaryNVR() {}

        public DictionaryNVR(IEnumerable<KeyValuePair<string, object>> dict)
        {
            if (this == null) throw new NullReferenceException();
            if (dict == null) throw new NullReferenceException();

            foreach (KeyValuePair<string, object> kvp in dict)
            {
                Add(kvp.Key, kvp.Value);
            }
        }

        public DictionaryNVR(IDictionary<string, object> dict)
        {
            if (this == null) throw new NullReferenceException();
            if (dict == null) throw new NullReferenceException();

            foreach (KeyValuePair<string, object> kvp in dict)
            {
                Add(kvp.Key, kvp.Value);
            }
        }

        public new object this[string key]
        {
            get
            {
                if (!ContainsKey(key.Trim())) throw new KeyNotFoundException("Missing key '" + key.Trim() + "'in dictionary:\n\n" + this);

                object value;
                TryGetValue(key, out value);
                return value;
            }
            set
            {
                Remove(key);
                Add(key, value);
            }
        }

        public void Sort()
        {
            if (this == null) throw new NullReferenceException();

            Dictionary<string, object> temp = this.OrderBy(entry => entry.Key).ToDictionary(pair => pair.Key, pair => pair.Value);

            Clear();

            foreach (KeyValuePair<string, object> kvp in temp)
            {
                Add(kvp.Key, kvp.Value);
            }
        }

        public override string ToString()
        {
            if (this == null) throw new NullReferenceException();

            string result = "";
            foreach (KeyValuePair<string, object> keyValuePair in this.OrderBy(entry => entry.Key).ToDictionary(pair => pair.Key, pair => pair.Value))
            {
                string count = "";
                object value = keyValuePair.Value;

                if (value != null)
                {
                    Type t = value.GetType();
                    try
                    {
                        count = " [Count=" + t.GetProperty("Count").GetValue(keyValuePair.Value, new object[0]) + "]";
                    }
                    catch {}
                }

                result += "'" + keyValuePair.Key + "': " + ( value ?? "NULL" ) + "" + count + "\n";
            }
            return result + "";
        }

        public object TryGetValue(string key)
        {
            return ContainsKey(key) ? this[key] : null;
        }
    }
}