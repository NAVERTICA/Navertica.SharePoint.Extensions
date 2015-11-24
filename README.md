NAVERTICA SharePoint Extensions (SSOM)
======================================

A selection of extension methods for SharePoint 2013 server-side object model, trying to make SharePoint development less of a pain particularly when used with dynamic languages. 

To see what we're doing with DLR (Dynamic Language Runtime) languages like IronPython, check out https://github.com/NAVERTICA/SPTools - with it you can use scripting to handle all kinds of SharePoint development and features, have all configurations available in a single place and functionalities activated in several places at once  with a simple URL routing with wildcards.

Feedback, ideas, patches welcome. 

Less scaffolding, less coupling
==========================
We tried to solve many of the common tasks routinely required of SharePoint programmers once and for all with convenient extensions 
on SharePoint objects, and take care of some of the many possible inconsistencies in the API behind the scenes.

We also tried to solve some of the coupling issues - you can't just pass anything SPRequest related (SPListItem etc.) to another
thread, so we're often working with our own classes like WebListId and WebListItemId, which contain the IDs of the required objects
and method to transparently access them, without being bound to specific SPRequests.

Another, much more daring, goal, which we have not reached by far and with the decline of SSOM probably never will, was to basically replace calls to SharePoint API with calls to our own extensions and 
thus decouple our business code from SharePoint API, which should've made it easier to replace the underlying SSOM with
something else in the future.

We make it easy to 
- open webs, lists, process items, without having to care about opening and disposing SPRequest connected objects and without pondering "do I have to use a Guid, or was it internal name? Oh, in this case it was display name, damn these inconsistencies..."
- copy or move items in both document libraries and custom lists - including metadata, attachments...
- look up content types by both name and ID
- check if fields with internal name exist in a list or content type
- find lists by type or content type
- do a lot of other things you probably already needed or will need to do shortly, if you're a SharePoint developer.

**Get and Set SPListItem extensions** can read (and write, where possible) normalized values of SPListItems, SPListItemVersions,
SPItemEventProperties (for Event Receivers with AfterProperties). All the returned values look the same,
whichever object they come from, and writing the values also works the same independent of whether the
underlying object is an SPListItem or AfterProperties of an item in receiver.

With Get and Set you can forget about having to do stuff like this:
```csharp
var myUsers = new SPFieldUserValueCollection();
foreach (string usr in MyUserList) {
	var usrValue = web.EnsureUser(usr);
	myUsers.Add(new SPFieldUserValue(web, usrValue.ID, usrValue.Name));
}
item["MyUserField"] = myUsers;
```
And instead just do this:
```csharp
item.Set("MyUserField", MyUserList);
```
Similarly, you could pass an enumerable with logins, names, emails, IDs, even mixed, or just put it all in a long semicolon-separated
string - it would still get processed just fine.

Using delegates to process (not only) items
---------------------------------------
It's not necessary to think about whether to open and dispose of SPWebs and other SPRequest-bound objects,
or having to open a few items at a time and paginate.

Now processing items returned by a CAML query can look like this:
```csharp
var titleCollection = list.ProcessItems(item => item["Title"], filterSPQuery);
```

Need to process every file in a folder and it's subfolders, and perhaps filter them by a complex query? You can forget about manually throttling the query, and just do this:
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
var titleCollection = item.ProcessLookupItems("LookupInternalName", itemInLookup => itemInLookup["Title"]);
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

Copyright (C) 2015 NAVERTICA a.s. http://www.navertica.com 
This program is free software; you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation; either version 2 of the License, or
(at your option) any later version.
