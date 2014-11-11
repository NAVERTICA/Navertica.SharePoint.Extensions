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
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint;

namespace Navertica.SharePoint.Extensions
{
    public static class SPContentTypeExtensions
    {
        /// <summary>
        /// Checks if content type collection contains content type with specified contentTypeId or name as string.
        /// </summary>
        /// <param name="contentTypeCollection"></param>
        /// <param name="identification">will be trimmed</param>
        /// <returns></returns>
        public static bool Contains(this SPContentTypeCollection contentTypeCollection, string identification)
        {
            if (contentTypeCollection == null) throw new ArgumentNullException("contentTypeCollection");
            if (identification == null) throw new ArgumentNullException("identification");
            if (identification.Trim() == string.Empty) throw new ArgumentException("empty identification");

            try
            {
                // in case identification is ContentTypeId
                return Contains(contentTypeCollection, new SPContentTypeId(identification.Trim()));
            }
                // when identification doesn't look like SPContentTypeId
            catch (ArgumentException) {}

            // in case identification is name
            List<string> list = contentTypeCollection.Cast<SPContentType>().Select(ct => ct.Name).ToList();
            return list.Any(name => name == identification.Trim());
        }

        /// <summary>
        /// Checks if content type collection contains content with given contentTypeId
        /// </summary>
        /// <param name="contentTypeCollection"></param>
        /// <param name="contentTypeId"></param>
        /// <returns></returns>
        public static bool Contains(this SPContentTypeCollection contentTypeCollection, SPContentTypeId contentTypeId)
        {
            if (contentTypeCollection == null) throw new ArgumentNullException("contentTypeCollection");
            if (contentTypeId == null) throw new ArgumentNullException("contentTypeId");

            List<string> list = contentTypeCollection.Cast<SPContentType>().Select(ct => ct.Id.ToString()).ToList();
            return list.Any(ctId => ctId.StartsWith(contentTypeId.ToString())); // CHECK - does adding 0x0108 to a list mean the list will accept 0x01 also?
        }

        /// <summary>
        /// Checks whether the list contains all the fields with internal names passed in intFieldNames
        /// </summary>
        /// <param name="contentType"></param>
        /// <param name="intFieldNames">internal names the list should contain</param>
        /// <returns></returns>
        public static bool ContainsFieldIntName(this SPContentType contentType, IEnumerable<string> intFieldNames)
        {
            if (contentType == null) throw new ArgumentNullException("contentType");
            if (intFieldNames == null) throw new ArgumentNullException("intFieldNames");

            foreach (string fieldname in intFieldNames)
            {
                try
                {
                    contentType.Fields.GetFieldByInternalName(fieldname.Trim());
                }
                catch
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Checks whether the content type contains a field with internal name
        /// </summary>
        /// <param name="contentType"></param>
        /// <param name="fieldName">field internal name - will be trimmed</param>
        /// <returns></returns>
        public static bool ContainsFieldIntName(this SPContentType contentType, string fieldName)
        {
            if (contentType == null) throw new ArgumentNullException("contentType");
            if (fieldName == null) throw new ArgumentNullException("fieldName");
            if (fieldName.Trim() == string.Empty) throw new ArgumentException("fieldName is empty string");

            try
            {
                contentType.Fields.GetFieldByInternalName(fieldName);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Get content type with specified contentTypeId
        /// </summary>
        /// <param name="contentTypeCollection"></param>
        /// <param name="contentTypeId">content type id as string or name - will be trimmed</param>
        /// <returns></returns>
        public static SPContentType GetContentType(this SPContentTypeCollection contentTypeCollection, SPContentTypeId contentTypeId)
        {
            if (contentTypeCollection == null) throw new ArgumentNullException("contentTypeCollection");
            if (contentTypeId == null) throw new ArgumentNullException("contentTypeId");

            return GetContentType(contentTypeCollection, contentTypeId.ToString());
        }

        /// <summary>
        /// Get content type with specified contentTypeId or name from collection
        /// </summary>
        /// <param name="contentTypeCollection"></param>
        /// <param name="identification">content type id as string or name - will be trimmed</param>
        /// <returns></returns>
        public static SPContentType GetContentType(this SPContentTypeCollection contentTypeCollection, string identification)
        {
            if (contentTypeCollection == null) throw new ArgumentNullException("contentTypeCollection");
            if (identification == null) throw new ArgumentNullException("identification");
            if (identification.Trim() == string.Empty) throw new ArgumentException("identification is empty string");

            try
            {
                SPContentTypeId searchCtId = new SPContentTypeId(identification);
                //if the contentTypeCollection is from SPList we are looking for their parent, TODO mozna dat dalsi parametr bool at cekuje i rodice to samy k funckcim contains
                return contentTypeCollection.Cast<SPContentType>().Single(ct => ct.Id == searchCtId || ct.Parent.Id == searchCtId); //identification is ContentTypeId
            }
            catch
            {
                return contentTypeCollection.Cast<SPContentType>().SingleOrDefault(ct => ct.Name == identification.Trim()); //identification is ContentType name
            }
        }
    }
}