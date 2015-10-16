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
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.SharePoint;

namespace Navertica.SharePoint.Extensions
{
    /// <summary>
    /// Navertica SharePoint Tools
    /// </summary>
    public static class StringExtensions
    {
        /// <summary>
        /// Returns true if given string is contained in the enumerable. One of element in enumerable is equal to str
        /// </summary>
        /// <param name="str"></param>
        /// <param name="enumerable"> </param>
        /// <returns></returns>
        public static bool ContainedIn(this IEnumerable<char> str, IEnumerable<string> enumerable)
        {
            if (str == null) throw new ArgumentNullException("str");
            if (enumerable == null) throw new ArgumentNullException("enumerable");

            return enumerable.Contains((string) str);
        }

        /// <summary>
        /// Returns true if any element in the enumerable contains given string
        /// </summary>
        /// <param name="str"></param>
        /// <param name="enumerable"></param>
        /// <returns></returns>
        public static bool ContainedInAnyString(this IEnumerable<char> str, IEnumerable<string> enumerable)
        {
            if (str == null) throw new ArgumentNullException("str");
            if (enumerable == null) throw new ArgumentNullException("enumerable");

            return enumerable.Any(s => s.Contains((string) str));
        }

        /// <summary>
        /// Checks if str contains any element in enumerable
        /// </summary>
        /// <param name="str"></param>
        /// <param name="enumerable"></param>
        /// <returns></returns>
        public static bool ContainsAny(this IEnumerable<char> str, IEnumerable<string> enumerable)
        {
            if (str == null) throw new ArgumentNullException("str");
            if (enumerable == null) throw new ArgumentNullException("enumerable");

            bool res = false;

            foreach (string s in enumerable)
            {
                if (( (string) str ).Contains(s)) res = true;
            }

            return res;
        }

        // reflection from Microsoft.JScript.dll
        public static string EscapeJavaScript(this string str)
        {
            string str2 = "0123456789ABCDEF";
            int length = str.Length;
            StringBuilder builder = new StringBuilder(length * 2);
            int num3 = -1;
            while (++num3 < length)
            {
                char ch = str[num3];
                int num2 = ch;
                if (( ( ( 0x41 > num2 ) || ( num2 > 90 ) ) &&
                      ( ( 0x61 > num2 ) || ( num2 > 0x7a ) ) ) &&
                    ( ( 0x30 > num2 ) || ( num2 > 0x39 ) ))
                {
                    switch (ch)
                    {
                        case '@':
                        case '*':
                        case '_':
                        case '+':
                        case '-':
                        case '.':
                        case '/':
                            goto Label_0125;
                    }
                    builder.Append('%');
                    if (num2 < 0x100)
                    {
                        builder.Append(str2[num2 / 0x10]);
                        ch = str2[num2 % 0x10];
                    }
                    else
                    {
                        builder.Append('u');
                        builder.Append(str2[( num2 >> 12 ) % 0x10]);
                        builder.Append(str2[( num2 >> 8 ) % 0x10]);
                        builder.Append(str2[( num2 >> 4 ) % 0x10]);
                        ch = str2[num2 % 0x10];
                    }
                }
                Label_0125:
                builder.Append(ch);
            }
            return builder.ToString();
        }

        /// <summary>
        /// Returns true if the given string is equal to any element in the enumerable 
        /// </summary>
        /// <param name="enumerable"></param>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool EqualAny(this IEnumerable<char> str, IEnumerable<string> enumerable)
        {
            if (str == null) throw new ArgumentNullException("str");
            if (enumerable == null) throw new ArgumentNullException("enumerable");

            return enumerable.Any(s => ( (string) str ) == s);
        }

        /// <summary>
        /// Creates a dict of lookup values with id as key "12;#Value;#;13;#OtherValue" returns dict { {12, "Value"}, {13, "OtherValue"} }
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static Dictionary<int, string> GetLookupDict(this IEnumerable<char> str)
        {
            if (str == null) throw new ArgumentNullException("str");

            return new SPFieldLookupValueCollection((string) str).ToDictionary(val => val.LookupId, val => val.LookupValue);
        }

        /// <summary>
        /// Returns the id part of a lookup string "12;#Value" returns 12
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static int GetLookupIndex(this IEnumerable<char> str)
        {
            if (str == null) throw new ArgumentNullException("str");
            if ((string) str == string.Empty) return -1;

            return GetLookupIndexes(str)[0];
        }

        /// <summary>
        /// Returns the id parts of a lookup string "12;#Value;#;13;#OtherValue" returns [12, 13]
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static int[] GetLookupIndexes(this IEnumerable<char> str)
        {
            if (str == null) throw new ArgumentNullException("str");

            string s = str.ToString();
            List<int> results = new List<int>();

            if (s.Trim() != string.Empty)
            {
                string[] vals = s.Split(";#");
                for (int i = 0; i < vals.Length; i = i + 2)
                {
                    int index;
                    Int32.TryParse(vals[i].Replace(";", ""), out index);
                    results.Add(index);
                }
            }

            return results.ToArray<int>();
        }

        /// <summary>
        /// Clears initial part of lookup/calculated strings "12;#Value" -> "Value"
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string GetLookupValue(this IEnumerable<char> str)
        {
            if (str == null) throw new ArgumentNullException("str");

            string lookupStr = str.ToString();

            if (lookupStr.Contains(";#"))
            {
                try
                {
                    return new SPFieldLookupValue(lookupStr).LookupValue;
                }
                catch //(ArgumentException err) // abychom mu mohli poslat i vypocitane pole, ktere na zacatku nema ID
                {
                    return lookupStr.Substring(lookupStr.IndexOf(";#", StringComparison.InvariantCulture) + 2);
                }
            }
            return lookupStr;
        }

        /// <summary>
        /// Clears initial part of lookup/calculated strings "12;#Value;#;13;#OtherValue" returns ["Value", "OtherValue"]
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string[] GetLookupValues(this IEnumerable<char> str)
        {
            if (str == null) throw new ArgumentNullException("str");

            string s = str.ToString();
            List<string> results = new List<string>();

            if (s.Trim() != string.Empty)
            {
                string[] vals = s.Split(";#");
                for (int i = 1; i < vals.Length; i = i + 2)
                {
                    results.Add(vals[i]);
                }
            }

            return results.ToArray<string>();
        }

        /// <summary>
        /// Returns an MD5 sum of a string
        /// </summary>
        /// <param name="str"></param>
        /// <returns>hash</returns>
        public static string GetMd5Sum(this IEnumerable<char> str)
        {
            if (str == null) throw new ArgumentNullException("str");

            StringBuilder sb = new StringBuilder();
            MD5CryptoServiceProvider md5CryptoServiceProvider = new MD5CryptoServiceProvider();
            foreach (byte b in md5CryptoServiceProvider.ComputeHash(Encoding.UTF8.GetBytes(str.ToArray()))) sb.Append(b.ToString("x2").ToLower());

            return sb.ToString();
        }

        /// <summary>
        /// Gets the paremeters from url string to dictionary. Keys are in lowercase
        /// </summary>
        /// <param name="url"></param>
        /// <returns>Filled Dictionary or empty. Never null</returns>
        public static DictionaryNVR GetParametersFromUrl(this string url)
        {
            if (url == null) throw new ArgumentNullException("url");

            DictionaryNVR dict = new DictionaryNVR();

            if (!url.Contains("?")) return dict;

            try
            {
                string queryString = url.Split('?').Last();
                string[] parameters = queryString.Split('&');

                foreach (string parameter in parameters)
                {
                    string[] values = parameter.Split('=');

                    dict.Add(values[0].ToLowerInvariant(), values[1]);
                }

                return dict;
            }
            catch (Exception)
            {
                return dict;
            }
        }

        /// <summary>
        /// Validates an email address
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool IsEmail(this IEnumerable<char> str)
        {
            if (string.IsNullOrEmpty((string) str)) return false;

            const string emailRegex = @"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}" +
                                      @"\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" +
                                      @".)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$";

            return new Regex(emailRegex).IsMatch((string) str);
        }

        /// <summary>
        /// Converts items to strings and joins them using joinWith
        /// </summary>
        /// <param name="enumerable"></param>
        /// <param name="joinWith"></param>
        /// <returns></returns>
        public static string JoinStrings(this IEnumerable enumerable, string joinWith = "\n")
        {
            if (enumerable == null) throw new ArgumentNullException("enumerable");
            if (joinWith == null) throw new ArgumentNullException("joinWith");

            return String.Join(joinWith, ( enumerable.Cast<object>().Select(o => ( o ?? "" ).ToString()).ToArray() ));
        }

        /// <summary>
        /// Reduce string to shorter preview which is optionally ended by some string (...).
        /// </summary>
        /// <param name="str">string to reduce</param>
        /// <param name="count">Length of returned string including endings.</param>
        /// <param name="endings">optional edings of reduced text</param>
        /// <example>
        /// string description = "This is very long description of something";
        /// string preview = description.Reduce(20,"...");
        /// produce -> "This is very long..."
        /// </example>
        /// <returns></returns>
        public static string Reduce(this IEnumerable<char> str, int count, string endings)
        {
            if (str == null) throw new ArgumentNullException("str");
            if (endings == null) endings = String.Empty;
            if (count < endings.Length) throw new Exception("Failed to reduce to less then endings length.");

            string s = str.ToString();

            int len = s.Length + endings.Length;
            if (count > s.Length) return s; //it's too short to reduce
            s = s.Substring(0, s.Length - len + count);
            s += endings;

            return s;
        }

        public static string RemoveAllWhiteSpaces(this IEnumerable<char> str)
        {
            if (str == null) throw new ArgumentNullException("str");

            var sb = new StringBuilder();
            foreach (char c in str.Where(c => !char.IsWhiteSpace(c)))
            {
                sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Remove characters with ACII values 0-31 and 127-159
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string RemoveControlCharacters(this IEnumerable<char> str)
        {
            if (str == null) throw new ArgumentNullException("str");

            return str.Where(character => !char.IsControl(character))
                .Aggregate(new StringBuilder(), (builder, character) => builder.Append(character))
                .ToString();
        }

        public static string RemoveDiacritics(this IEnumerable<char> str)
        {
            if (str == null) throw new ArgumentNullException("str");

            StringBuilder sb = new StringBuilder();

            string s = ( (string) str ).Normalize(NormalizationForm.FormD); // oddeleni znaku od modifikatoru (hacku, carek, atd.) 

            foreach (char t in s.Where(t => CharUnicodeInfo.GetUnicodeCategory(t) != UnicodeCategory.NonSpacingMark))
            {
                sb.Append(t);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Removes from string the following characters '\a' '\b' '\t' '\n' '\v' '\f' '\r'
        /// </summary>
        /// <param name="str"> </param>
        /// <returns></returns>
        public static string RemoveEscapedCharacters(this IEnumerable<char> str)
        {
            if (str == null) throw new ArgumentNullException("str");

            string result = str.ToString();
            for (int i = 7; i < 14; i++)
            {
                result = result.Replace(( (char) i ).ToString(CultureInfo.InvariantCulture), "");
            }

            return result;
        }

        public static string RemoveFirstAndLastCharacters(this IEnumerable<char> str)
        {
            if (str == null) throw new ArgumentNullException("str");

            string tmp = str.ToString();
            tmp = tmp.Remove(0, 1);
            tmp = tmp.Remove(tmp.Length - 1, 1);
            return tmp;
        }

        /// <summary>
        /// remove white space, not line end
        /// Useful when parsing user input such phone,
        /// price int.Parse("1 000 000".RemoveSpaces(),.....
        /// </summary>
        /// <param name="str"></param>
        public static string RemoveSpaces(this IEnumerable<char> str)
        {
            if (str == null) throw new ArgumentNullException("str");

            return ( (string) str ).Replace(" ", "");
        }

        /// <summary>
        /// Replaces all occurrences of oldValue by newValue in all string in the enumerable.
        /// </summary>
        /// <param name="enumerable"></param>
        /// <param name="oldValue"></param>
        /// <param name="newValue"></param>
        /// <returns></returns>
        public static IEnumerable<string> Replace(this IEnumerable<string> enumerable, string oldValue, string newValue)
        {
            return enumerable.Select(s => s.Replace(oldValue, newValue));
        }

        /// <summary>
        /// Alternate String.Split, works with a single string
        /// </summary>
        /// <param name="delimiter"></param>
        /// <param name="str"></param>
        /// <param name="options"></param>
        /// <returns></returns>
        public static string[] Split(this IEnumerable<char> str, string delimiter, StringSplitOptions options = StringSplitOptions.None)
        {
            if (str == null) throw new ArgumentNullException("str");
            if (delimiter == null) throw new ArgumentNullException("delimiter");

            return ( (string) str ).Split(new[] { delimiter }, options);
        }

        /// <summary>
        /// Takes chars from charsToSplitBy, and uses the first one contained in toSplit to split the string.
        /// It leaves out empty strings and trims every item.        
        /// </summary>
        /// <param name="str"></param>
        /// <param name="charsToSplitBy"></param>
        /// <returns></returns>
        public static string[] SplitByChars(this IEnumerable<char> str, string charsToSplitBy)
        {
            if (str == null) throw new ArgumentNullException("str");
            if (charsToSplitBy == null) throw new ArgumentNullException("charsToSplitBy");

            string[] result = null;
            string s = str.ToString();

            foreach (char ch in charsToSplitBy)
            {
                if (s.Contains(ch))
                {
                    result = s.Split(new[] { ch }, StringSplitOptions.RemoveEmptyEntries);
                    break;
                }
            }

            if (result == null)
            {
                result = new[] { s };
            }

            return result.Trim().ToArray();
        }

        /// <summary>
        /// Takes chars from ";,| \n", and uses the first one contained in toSplit to split the string.
        /// It leaves out empty strings and trims every item.        
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string[] SplitByCharsDefault(this IEnumerable<char> str)
        {
            if (str == null) throw new ArgumentNullException("str");

            return SplitByChars(str, ";,| \n"); // ne dvojtecka, protoze tu pouzivame na rozdeleni guidu
        }

        /// <summary>
        /// Splits string to smaller pieces of given length
        /// </summary>
        /// <param name="str"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public static string[] SplitByLength(this IEnumerable<char> str, int length)
        {
            if (str == null) throw new ArgumentNullException("str");

            string s = str.ToString();
            int len = s.Length;
            List<string> myArray = new List<string>();

            while (len > length)
            {
                string newString = s.Substring(0, length);
                s = s.Substring(length);
                myArray.Add(newString);
                len = s.Length;
            }

            myArray.Add(s);

            return myArray.ToArray();
        }

        /// <summary>
        /// Checks if str starts with any element in enumerable
        /// </summary>
        /// <param name="str"></param>
        /// <param name="enumerable"></param>
        /// <returns></returns>
        public static bool StartsWithAny(this IEnumerable<char> str, IEnumerable<string> enumerable)
        {
            if (str == null) throw new ArgumentNullException("str");
            if (enumerable == null) throw new ArgumentNullException("enumerable");

            return enumerable.Any(s => ( (string) str ).StartsWith(s));
        }

        /// <summary>
        /// Strips HTML tags from string
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string StripHTML(this IEnumerable<char> str)
        {
            if (str == null) throw new ArgumentNullException("str");

            return Regex.Replace((string) str, "<.*?>", string.Empty);
        }

        public static IEnumerable<string> ToLowerInvariant(this IEnumerable<string> enumerable)
        {
            if (enumerable == null) throw new ArgumentNullException("enumerable");

            return enumerable.Select(ret => ret.ToLowerInvariant());
        }

        public static IEnumerable<string> ToUpperInvariant(this IEnumerable<string> enumerable)
        {
            if (enumerable == null) throw new ArgumentNullException("enumerable");

            return enumerable.Select(ret => ret.ToUpperInvariant());
        }

        /// <summary>
        /// Converts UTF8 string to byte array
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static byte[] ToUtf8ByteArray(this IEnumerable<char> str)
        {
            if (str == null) throw new ArgumentNullException("str");

            UTF8Encoding encoding = new UTF8Encoding();
            return encoding.GetBytes(str.ToString());
        }

        public static IEnumerable<string> Trim(this IEnumerable<string> enumerable)
        {
            if (enumerable == null) throw new ArgumentNullException("enumerable");

            return enumerable.Select(ret => ret.Trim());
        }
    }
}