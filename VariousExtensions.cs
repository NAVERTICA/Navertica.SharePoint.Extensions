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
using System.Globalization;
using System.IO;
using System.Linq;
using Microsoft.SharePoint;

namespace Navertica.SharePoint.Extensions
{
    /// <summary>
    /// Navertica SharePoint Tools
    /// </summary>
    public static class ExtensionsTools
    {
        public static void AddViewFields(this SPView view, IEnumerable<string> fields, bool deleteCurrentViewFields = false)
        {
            if (view == null) throw new ArgumentNullException("view");

            if (deleteCurrentViewFields)
            {
                view.ViewFields.DeleteAll();
            }

            foreach (string fldIntName in fields)
            {
                if (!view.ViewFields.Exists(fldIntName))
                {
                    view.ViewFields.Add(fldIntName);
                }
            }

            view.Update();
        }

        public static string GetProperties(this object obj)
        {
            if (obj == null) throw new ArgumentNullException("obj");

            return obj.GetType().GetProperties().OrderBy(i => i.Name).Select(i => i.Name + " [" + i.PropertyType.Name + "]").JoinStrings();
        }

        public static DictionaryNVR GetValueForAllCultures(this SPUserResource resource)
        {
            SPWeb web = resource.Parent.GetType() == typeof (SPWeb) ? (SPWeb) resource.Parent : ( (SPList) resource.Parent ).ParentWeb;

            DictionaryNVR values = new DictionaryNVR();
            foreach (CultureInfo info in web.RegionalSettings.InstalledLanguages.Cast<SPLanguage>().Select(l => CultureInfo.GetCultureInfo(l.LCID)))
            {
                values[info.LCID.ToString(CultureInfo.InvariantCulture)] = resource.GetValueForUICulture(info);
            }

            return values;
        }

        public static bool IsEmpty(this Guid guid)
        {
            return guid == Guid.Empty;
        }

        /// <summary>
        /// Returns true if datetime is in interval
        /// </summary>
        /// <param name="dateTime"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public static bool IsInInterval(this DateTime dateTime, DateTime start, DateTime end)
        {
            if (dateTime == null) throw new ArgumentNullException("dateTime");

            return start.CompareTo(dateTime) < 0 && end.CompareTo(dateTime) > 0;
        }

        /// <summary>
        /// Reads data from a stream until the end is reached. The
        /// data is returned as a byte array. An IOException is
        /// thrown if any of the underlying IO calls fail.
        /// </summary>
        /// <param name="stream">The stream to read data from</param>
        public static byte[] ReadFully(this Stream stream)
        {
            if (stream == null) throw new ArgumentNullException();

            byte[] buffer = new byte[32768];

            using (MemoryStream ms = new MemoryStream())
            {
                while (true)
                {
                    int read = stream.Read(buffer, 0, buffer.Length);
                    if (read <= 0)
                        return ms.ToArray();
                    ms.Write(buffer, 0, read);
                }
            }
        }

        /// <summary>
        /// Replaces invalid characters in a file name
        /// </summary>
        /// <param name="fname"></param>
        /// <param name="replaceWith"></param>
        /// <returns></returns>
        public static string ReplaceInvalidFileNameChars(this string fname, char replaceWith)
        {
            // pokud mame v datech _UID, je to string obsahujici strednikem oddelena jmena parametru, ktera tvori primarni klic 
            // a ktera tedy predradime pred jmeno souboru, aby se soubory se stejnym jmenem pro ruzne polozky neprepisovaly
            List<char> invalidChars = new List<char>(Path.GetInvalidFileNameChars()) { '#', '%', '&', '*', ':', '<', '>', '?', '/', '{', '|', '}' };

            return invalidChars.Aggregate(fname, (current, c) => current.Replace(c, replaceWith)).Replace("..", replaceWith.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Reads data from a stream until the end is reached. The
        /// data is returned as a byte array. An IOException is
        /// thrown if any of the underlying IO calls fail.
        /// http://www.yoda.arachsys.com/csharp/readbinary.html
        /// </summary>
        /// <param name="stream">The stream to read data from</param>
        /// 
        public static byte[] StreamToByteArray(this Stream stream)
        {
            if (stream == null) throw new ArgumentNullException();

            byte[] buffer = new byte[32768];
            using (MemoryStream ms = new MemoryStream())
            {
                while (true)
                {
                    int read = stream.Read(buffer, 0, buffer.Length);
                    if (read <= 0) return ms.ToArray();
                    ms.Write(buffer, 0, read);
                }
            }
        }

        #region Conversion functions

        /// <summary>
        /// Returns false if object is null else true. If not null tries parse value to boolean
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public static bool ToBool(this object val)
        {
            if (val is bool) return (bool) val;

            string value = ( val ?? "" ).ToString().Trim();

            bool result = false;
            if (string.IsNullOrEmpty(value)) return false;

            try
            {
                result = Convert.ToBoolean(val);
            }
            catch
            {
                try
                {
                    int boolVal = int.Parse(value);
                    result = Convert.ToBoolean(boolVal);
                }
                catch
                {
                    if (value.ToLowerInvariant().EqualAny(new[] { "on", "yes", "ano" })) result = true;
                }
            }
            return result;
        }

        public static int ToInt(this object value)
        {
            if (value is int) return (int) value;

            try
            {
                return Convert.ToInt32(value);
            }
            catch
            {
                return Int32.MinValue;
            }
        }

        public static double ToDouble(this object value)
        {
            try
            {
                return Convert.ToDouble(value);
            }
            catch
            {
                return Double.NaN;
            }
        }

        #endregion

        public static string ToStringISO(this DateTime dateTime, CultureInfo culture = null)
        {
            if (dateTime == null) throw new ArgumentNullException("dateTime");

            if (culture == null) culture = CultureInfo.InvariantCulture;

            return dateTime.ToString("s", culture);
        }

        public static string ToStringLocalized(this DateTime dateTime, bool incudeTime = false, int lang = -1)
        {
            if (dateTime == null) throw new ArgumentNullException("dateTime");

            CultureInfo culture = CultureInfo.CurrentUICulture;
            DateTimeFormatInfo format = culture.DateTimeFormat;
            if (lang > 0)
            {
                culture = CultureInfo.GetCultureInfo(lang);
                format = culture.DateTimeFormat;
            }
            string result = dateTime.ToString(format.ShortDatePattern, culture);

            if (incudeTime) result += " " + dateTime.ToString(format.ShortTimePattern);
            return result;
        }
    }
}