NAVERTICA SharePoint Extensions (SSOM)
======================================

A selection of extension methods for SharePoint 2013 server-side object model.


Feedback welcome.

Less scaffolding
=================
We tried to solve many of the common tasks routinely required of SharePoint programmers once and for all
with convenient extensions on SharePoint objects, 
and take care of some of the many possible inconsistencies in the API behind the scenes.

Opening a web or list? 
```csharp
SPWeb web = site.OpenW(Url_Name_Guid_AsString, false /* don't throw exception */);
if (web == null) logError("Couldn't open web " + UrlOrNameOrGuidAsString);
SPList list = web.OpenList(Url_Name_InternalName_Guid_AsString, false);
if (list == null) logError("Couldn't open list " + UrlOrNameOrInternalNameOrGuidAsString);
```

TODO - Get and Set extensions, to read (and write, where possible) normalized values of SPListItems, SPListItemVersions,
SPItemEventProperties (for Event Receivers with AfterProperties). All the returned values look the same,
whichever object they come from, and writing the values also works the same independent of whether the
underlying object is an SPListItem or AfterProperties of an item in receiver.

Using delegates to process (not only) items
---------------------------------------
It's not necessary to think about whether to open and dispose of SPWebs and other SPRequest-bound objects,
or having to open a few items at a time and paginate.

Now processing items returned by a CAML query can look like this:
```csharp
var titleCollection = list.ProcessItems(item => item["Title"], filterSPQuery);
```

Need to process every file in a folder and it's subfolders, and perhaps filter them by a complex query?
```csharp
var resultingNumberFieldValuesCollection = 
	folder.ProcessItems(
		delegate(SPListItem item) { 
				item["NumberField"] = item["NumberField"] * 2; 
				item.Update(); 
				return item["NumberField"]; 
		}, optionalSPQuery, true);
```

Process all items in a lookup? 
```csharp
var titleCollection = item.ProcessLookupItems("LookupInternalName", item => item["Title"]);
```

Query Tools - strongly typed CAML query helper classes
---------------------------------------
Build queries incrementally
```csharp
Q query = new Q();
if (!string.IsNullOrEmpty(Value1))
{
	query.Add(new Q(QOp.Equal, QVal.Text, "FieldValue1", Value1.Trim()));
}
if (!string.IsNullOrEmpty(Value2))
{
	query.Add(new Q(QOp.BeginsWith, QVal.Text, "FieldValue2", Value2.Trim()));
}
SPListItemCollection col = list.GetItems(query.ToString());
```

Safely build complex queries, integrate with existing CAML queries and query fragments...
```csharp
Q q = 
	new Q(QJoin.And,
	  new Q(QJoin.And,
		new Q(QOp.Equal, valueType, config.ItemInternalName, item[config.ItemInternalName]),
		new Q(QJoin.Or,
			new Q(QJoin.And,	
				new Q(QOp.Greater, QVal.DateTime, config.EndDateInternalName, item[config.EndDateInternalName]),
				new Q(QOp.LesserOrEqual, QVal.DateTime, config.EndDateInternalName, item[config.EndDateInternalName]])
			),
			new Q(QJoin.And,	
				new Q(QOp.Lesser, QVal.DateTime, config.StartDateInternal, item[config.StartDateInternal]]),
				new Q(QOp.GreaterOrEqual, QVal.DateTime, config.StartDateInternal, item[config.StartDateInternal]])
			)
		)
	  ),
	  new Q("<DateRangesOverlap><FieldRef Name=\"EventDate\" /><FieldRef Name=\"EndDate\" /><FieldRef Name=\"RecurrenceID\" /><Value Type=\"DateTime\"><Today /></Value></DateRangesOverlap>"),
	  new Q(QOp.NotEqual, QVal.Counter, NVRField.ID, item.ID)
	  );
SPQuery query = new SPQuery
{
	ViewFields = ViewFlds,
	CalendarDate =
		SPUtility.CreateDateTimeFromISO8601DateTimeString(
			item[config.StartDateInternal].ToString()),
	ExpandRecurrence = true,
	Query = q.ToString(true)
};
SPListItemCollection col = list.GetItems(query);
```


