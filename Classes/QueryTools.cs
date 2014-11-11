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
using Microsoft.SharePoint.Utilities;
using Navertica.SharePoint.Extensions;

// ReSharper disable once CheckNamespace
namespace Navertica.SharePoint
{
    public abstract class QComponent {}

    #region enums

    public enum QOp
    {
        //http://msdn.microsoft.com/en-us/library/ms467521
        BeginsWith,
        Contains,
        //DateRangesOverlap
        Equal,
        GreaterOrEqual,
        Greater,
        //In,
        //Includes
        IsNotNull,
        IsNull,
        LesserOrEqual,
        Lesser,
        NotEqual,
    }

    public enum QVal
    {
        Boolean,
        Counter,
        Choice,
        Date,
        DateToday,
        Lookup,
        LookupId,
        Number,
        Text,
        User,
        URL,
        DateTime,
        DateTimeToday
    }

    public enum QJoin
    {
        And,
        Or
    }

    #endregion

    /// <summary>
    /// Wraps most functionality of CAML queries for more dependable handling. Example usage:
    /// Q query =
    ///     new Q(QJoin.And,
    ///         new Q(QJoin.Or,
    ///             new Q(QOp.Equal, QVal.Text, "Title", "TitleValue"),
    ///             new Q(QOp.IsNull, QVal.DateTime, "Expires", "")),
    ///         new Q(QOp.NotEqual, QVal.Boolean, "Disabled", "1"));
    /// </summary>
    public class Q : QComponent
    {
        private readonly QJoin _joinType = QJoin.And;
        private readonly Queue<QComponent> _queue = new Queue<QComponent>();
        private List<string> _orderBy = new List<string>();
        private List<string> _viewFields = new List<string>();

        #region Constructors

        public Q() {}

        public Q(QJoin joinType)
        {
            _joinType = joinType;
        }

        public Q(QOp op, QVal valueType, string intFieldName, object value, IEnumerable<string> orderBy = null, IEnumerable<string> viewFields = null)
        {
            _queue.Enqueue(new QLeaf(op, valueType, intFieldName, value));
            if (orderBy != null)
            {
                _orderBy = new List<string>(orderBy);
            }
            if (viewFields != null)
            {
                _viewFields = new List<string>(viewFields);
            }
        }

        public Q(params QComponent[] qList)
        {
            _queue = new Queue<QComponent>(qList);
        }

        public Q(QJoin joinType, params QComponent[] qList)
        {
            _joinType = joinType;
            _queue = new Queue<QComponent>(qList);
        }

        public Q(QJoin joinType, IEnumerable<string> orderBy, params QComponent[] qList)
        {
            _joinType = joinType;
            _queue = new Queue<QComponent>(qList);
            if (orderBy != null)
            {
                _orderBy = new List<string>(orderBy);
            }
        }

        public Q(QJoin joinType, IEnumerable<string> orderBy, IEnumerable<string> viewFields, params QComponent[] qList)
        {
            _joinType = joinType;
            _queue = new Queue<QComponent>(qList);
            if (orderBy != null)
            {
                _orderBy = new List<string>(orderBy);
            }
            if (viewFields != null)
            {
                _viewFields = new List<string>(viewFields);
            }
        }

        public Q(QJoin joinType, params string[] queryStrings)
        {
            _joinType = joinType;
            foreach (string q in queryStrings)
            {
                _queue.Enqueue(new QLeafString(q));
            }
        }

        public Q(params string[] queryStrings)
        {
            foreach (string q in queryStrings)
            {
                _queue.Enqueue(new QLeafString(q));
            }
        }

        public Q(IEnumerable<string> orderBy, params QComponent[] qList)
        {
            _queue = new Queue<QComponent>(qList);
            if (orderBy != null)
            {
                _orderBy = new List<string>(orderBy);
            }
        }

        public Q(IEnumerable<string> orderBy, IEnumerable<string> viewFields, params QComponent[] qList)
        {
            _queue = new Queue<QComponent>(qList);
            if (orderBy != null)
            {
                _orderBy = new List<string>(orderBy);
            }
            if (viewFields != null)
            {
                _viewFields = new List<string>(viewFields);
            }
        }

        #endregion

        private string GetOrderString()
        {
            string result = "";

            if (_orderBy != null && _orderBy.Count > 0)
            {
                result += "<OrderBy>";
                foreach (string ord in _orderBy)
                {
                    string fName = ord.Trim();
                    string descString = "";
                    if (fName.StartsWith("<") || fName.StartsWith(">"))
                    {
                        descString = fName.StartsWith(">") ? "Ascending='FALSE' " : "Ascending='TRUE' ";
                        fName = fName.Substring(1).Trim();
                    }
                    result += "<FieldRef Name='" + fName + "' " + descString + "/>";
                }
                result += "</OrderBy>";
            }
            return result;
        }

        private string GetViewFieldsString()
        {
            string result = "";

            if (_viewFields != null && _viewFields.Count > 0)
            {
                result += "<ViewFields>";
                result = _viewFields.Aggregate(result, (current, ord) => current + ( "<FieldRef Name='" + ord.Trim() + "' />" ));
                result += "</ViewFields>";
            }
            return result;
        }

        private string BuildQuery(Queue<QComponent> q, QJoin type)
        {
            if (q.Count > 0)
            {
                Queue<QComponent> items = new Queue<QComponent>(q);
                QComponent qcomp = items.Dequeue();

                string currentTerm = ToCamlString(qcomp);

                if (items.Count == 0) return currentTerm;

                return "<" + type + ">" + currentTerm + BuildQuery(items, type) + "</" + type + ">";
            }
            return string.Empty;
        }

        private string ToCamlString(QComponent q)
        {
            if (q.GetType() == typeof (QLeaf) || q.GetType() == typeof (QLeafString)) return q.ToString();

            Q qtemp = (Q) q;
            if (qtemp._orderBy != null) _orderBy.AddRange(qtemp._orderBy);
            if (qtemp._viewFields != null) _viewFields.AddRange(qtemp._viewFields);

            return BuildQuery(qtemp._queue, qtemp._joinType);
        }

        public void Add(params QComponent[] queries)
        {
            foreach (QComponent qry in queries)
            {
                _queue.Enqueue(qry);
            }
        }

        /// <summary>
        /// Builds the resulting CAML query string
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            string result = BuildQuery(_queue, _joinType);

            _orderBy = _orderBy.Distinct().Trim().ToList();
            _viewFields = _viewFields.Distinct().Trim().ToList();
            result += GetViewFieldsString();
            result += GetOrderString();

            return result;
        }

        /// <summary>
        /// Builds the resulting CAML query string and adds outer <Where></Where> 
        /// </summary>
        /// <returns></returns>
        public string ToString(bool surroundWithWhere)
        {
            if (surroundWithWhere)
            {
                return string.Format("<Where>{0}</Where>", this);
            }
            return ToString();
        }
    }

    #region internals

    internal class QLeafString : QComponent
    {
        public string Query = "";

        public QLeafString(string qString)
        {
            Query = qString;
        }

        public override string ToString()
        {
            return Query;
        }
    }

    internal class QLeaf : QComponent
    {
        public QOp Operation;
        public string FieldName;
        public QVal ValueType;
        public object Value;

        public QLeaf(QOp op, QVal valueType, string intFieldName, object value)
        {
            Operation = op;
            FieldName = intFieldName.Trim();
            ValueType = valueType;
            Value = value;
        }

        public override string ToString()
        {
            string opstr = OperationString();
            string valuetypestr = ValueTypeString();

            string valueXml = "";
            if (Value != null && Operation != QOp.IsNotNull && Operation != QOp.IsNull)
            {
                string val;
                if (ValueType == QVal.Boolean)
                {
                    val = Value.ToBool() ? "1" : "0";
                }
                else if (ValueType == QVal.DateTime || ValueType == QVal.Date)
                {
                    try
                    {
                        val = SPUtility.CreateISO8601DateTimeFromSystemDateTime(DateTime.Parse(Value.ToString()));
                    }
                    catch
                    {
                        val = Value.ToString();
                    }
                }
                else if (ValueType == QVal.DateTimeToday || ValueType == QVal.DateToday)
                {
                    val = "<Today";
                    if (Value != null && Value.ToString().Trim() != "")
                    {
                        val += " OffsetDays='" + Value + "'";
                    }
                    val += " />";
                }
                else
                {
                    val = Value.ToString();
                }

                if (ValueType == QVal.DateTime || ValueType == QVal.DateTimeToday)
                {
                    valuetypestr = valuetypestr + "' IncludeTimeValue='True";
                }

                val = val.Replace("\"", "&quot;");

                valueXml = string.Format("<Value Type='{0}'>{1}</Value>", valuetypestr, "<![CDATA[" + val + "]]>"); //value can contains some xml not supported characters... like < > etc...
            }

            return string.Format("<{0}>" + "<FieldRef Name='{1}'{2}/>{3}</{0}>",
                opstr,
                FieldName,
                ( ValueType == QVal.LookupId ? " LookupId='TRUE'" : string.Empty ),
                valueXml);
        }

        private string OperationString()
        {
            switch (Operation)
            {
                case QOp.BeginsWith:
                    return "BeginsWith";
                case QOp.Contains:
                    return "Contains";
                case QOp.Greater:
                    return "Gt";
                case QOp.GreaterOrEqual:
                    return "Geq";
                case QOp.Lesser:
                    return "Lt";
                case QOp.LesserOrEqual:
                    return "Leq";
                case QOp.NotEqual:
                    return "Neq";
                case QOp.IsNull:
                    return "IsNull";
                case QOp.IsNotNull:
                    return "IsNotNull";
                case QOp.Equal:
                    return "Eq";
                default:
                    return "Eq";
            }
        }

        private string ValueTypeString()
        {
            switch (ValueType)
            {
                case QVal.Counter:
                    return "Counter";
                case QVal.Choice:
                    return "Choice";
                case QVal.Date:
                case QVal.DateTime:
                case QVal.DateTimeToday:
                    return "DateTime";
                case QVal.Lookup:
                case QVal.LookupId:
                    return "Lookup";
                case QVal.Number:
                case QVal.Boolean:
                    return "Number";
                case QVal.URL:
                    return "URL"; //value musi byt URL bez site.Url
                case QVal.User:
                    return "User";
                case QVal.Text:
                    return "Text";
                default:
                    return "Text";
            }
        }

        #endregion
    }
}