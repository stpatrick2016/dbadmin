<%@ Language=VBScript %>
<!--#include file=scripts\inc_common.asp -->
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<head>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet" type="text/css">
<title>DBA:<%=langCaptionRecEdit%></title>
<%
	Const TableNameKey = "BBC017D1-0A13-4a9d-9A53-52A0CC6A7540"
	Const PKNameKey = "57AFDC29-37C8-48e1-96BE-12D7B79C1825"
	Const ActionNameKey = "714A51CC-797B-4ce5-99C7-81DB8721D68B"
	Const KeyNameKey = "0C61DD31-D805-4f58-A369-E4F33595FB86"
	Const NextPos = "B0802058-4FBB-4729-BA67-5CE30EDF3FC6"
	Dim Separator 
	Separator = vbTab
%>
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
<script language="javascript" type="text/javascript">
function onCancelEdit(table, page){
	location.href = "data.asp?table=" + table + "&page=" + page;
}
function onGoNext(str){
	document.getElementById('<%=NextPos%>').value = str;
}
function onAddNew(){
	onGoNext('next');
	document.getElementById('<%=ActionNameKey%>').value = '';
}
function onDelete(){
	var ret = confirm('<%=langSureToDeleteRecord%>');
	if(ret == true){
		onGoNext('delete');
	}
	return ret;
}
</script>
</head>
<body>
		
		
<%
	call DBA_WriteNavigation

	if not DBAE_DEBUG Then On Error Resume Next
	dim rec, sSQL, sTable, pk, key, fld, bIsEdit, sName, i
	dim action, strRedirect, oIndexes, page, varBookmark
	dim bHasPrev, bHasNext, bGoNext, bDoUpdate
	dim dba, PrimaryKeys, DefaultValue, item
	
	sTable = Request("table").Item
	pk = Split(Request("pk"), Separator)
	key = Split(Request("key"), Separator)
	action = Request(ActionNameKey)
	if Len(action) = 0 then action = Request("action").Item
	page = Request("page").Item
	
	bHasPrev = True
	bHasNext = True
	bGoNext = False
	bDoUpdate = False
	
	if action = "edit" then 
		bIsEdit = True
		sName = langUpdate
	else
		bIsEdit = False
		sName = langAdd
	end if
	set dba = new DBAdmin
	dba.Connect Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword)
	DBA_BeginNewTable sName & "&nbsp;" & langRecord, "", "", ""
%>
<p align=center><%=langAutoNumberNote%></p>
<p align=center><%=langRecEditNote%></p>
<a name="form"></a>
<%	

	if dba.HasError then DBA_WriteError dba.LastError
	if Request.Form(TableNameKey).Count > 0 then
		sTable = Request.Form(TableNameKey)
		pk = Split(Request.Form(PKNameKey), Separator)
		key = Split(Request.Form(KeyNameKey), Separator)
		action = Request(ActionNameKey)

		'retrieve string with primary keys names
		PrimaryKeys = ""
		for each i in dba.Tables.Item(sTable).Indexes.Items
			if i.IsPrimary then PrimaryKeys = PrimaryKeys & i.TargetField & Separator
		next
		if Right(PrimaryKeys, 1) = Separator then PrimaryKeys = Left(PrimaryKeys, Len(PrimaryKeys)-1)

		if InStr(1, Request.Form(NextPos), "update", vbTextCompare) > 0 or (Len(action) = 0 and UBound(key) = -1) then bDoUpdate = True
		if action = "edit" then
			sSQL = ""
			fld = ""
			for i=0 to UBound(pk)
				if Len(pk(i)) > 0 and Len(key(i)) > 0 then
					sSQL = sSQL & fld & " [" & pk(i) & "]=" & key(i)
					fld = " AND"
				end if
			Next
			bIsEdit = True
		else
			bIsEdit = False
		end if
		set rec = dba.Tables.Item(sTable).GetRawData(4, "[" & sTable & "]", False)
		if Len(sSQL) > 0 then rec.Filter = sSQL
		if rec.EOF or bIsEdit = False then 
			if bDoUpdate = True then rec.AddNew 
			bIsEdit = False
		else
			bIsEdit = True
		end if
		
		if bDoUpdate = True then
			for each fld in rec.Fields 
				if not fld.Properties("ISAUTOINCREMENT") and Len(fld.Name) > 0 then
					if Len(Request.Form(fld.Name)) = 0 then
						if bIsEdit then 
							if fld.Type = adWChar or fld.Type = adVarWChar or fld.Type = adLongVarWChar then
								fld.Value = "" 
							elseif fld.Type = adBoolean Then 
								fld.Value = False
							else 
								fld.Value = null
							End If
						end if
					elseif fld.Type = adDate then 
						fld.Value = CDate(Request.Form(fld.Name)) 
					else 
						fld.Value = Request.Form(fld.Name)
					end if
				end if
			Next
			rec.Update 
		end if
		varBookmark = rec.Bookmark
		if Err then
			call DBA_WriteError(Err.Description)
			rec.Close
		else
			rec.Filter = ""
			rec.Bookmark = varBookmark
			bGoNext = False
			Select Case Request.Form(NextPos)
				Case "next", "update_next"
					rec.MoveNext
					if not rec.EOF or action <> "edit" then bGoNext = True
				Case "prev", "update_prev"
					rec.MovePrevious
					if not rec.BOF Then bGoNext = true
				Case "first"
					rec.MoveFirst
					bGoNext = True
				Case "last"
					rec.MoveLast
					bGoNext = True
				Case "delete"
					rec.Delete
					rec.MovePrevious
					if not rec.EOF then rec.MoveNext 
					if not rec.EOF or not rec.BOF then bGoNext = True
			End Select
			
			if bGoNext then
				strRedirect = 	"recedit.asp?table=" & Server.URLEncode(sTable) &_
								"&pk=" & Server.URLEncode(Join(pk, Separator)) &_
								"&action=" & action &_
								"&page=" & page
				key = ""
				if action = "edit" then key = GetPKValues(Split(PrimaryKeys, Separator), rec, False)
				strRedirect =	strRedirect & "&key=" & Server.URLEncode(key)
				if bDoUpdate = True then strRedirect = strRedirect & "&message=" & Server.URLEncode(langRecordUpdated)
				strRedirect = strRedirect & "&d#form"
			else
				strRedirect = "data.asp?table=" & Server.URLEncode(sTable)
			end if
			rec.Close
			set rec = nothing
			Response.Redirect strRedirect
		end if
	end if
	
	if Len(PrimaryKeys) = 0 then
		'retrieve string with primary keys names
		PrimaryKeys = ""
		for each i in dba.Tables.Item(sTable).Indexes.Items
			if i.IsPrimary then PrimaryKeys = PrimaryKeys & i.TargetField & Separator
		next
		if Right(PrimaryKeys, 1) = Separator then PrimaryKeys = Left(PrimaryKeys, Len(PrimaryKeys)-1)
	end if
	
	sSQL = ""
	if action = "edit" then 
		fld = ""
		dim test
		if UBound(pk) = -1 then
			pk = Split(PrimaryKeys, Separator)
		else
			for i=0 to UBound(pk)
				if Len(pk(i)) > 0 and Len(key(i)) > 0 then
					sSQL = sSQL & fld & " [" & pk(i) & "]=" & key(i)
					fld = " AND"
				end if
			Next
		end if
	end if
	set rec = dba.Tables.Item(sTable).GetRawData(4, "[" & sTable & "]", True)
	if Len(Request.QueryString("filter_criteria").Item) > 0 Then
		rec.Find BuildFilter()
	Else
		if Len(sSQL) > 0 then rec.Filter = sSQL
		if action = "edit" and Len(sSQL) = 0 then key = GetPKValues(pk, rec, True)
	End If
	varBookmark = rec.Bookmark
	rec.Filter = ""
	rec.Bookmark = varBookmark
	
	if action = "edit" then
		rec.MoveNext
		if rec.EOF then bHasNext = False
		rec.MovePrevious
		rec.MovePrevious
		if rec.BOF then bHasPrev = False
		rec.Bookmark = varBookmark
	else
		bHasPrev = False
	end if
%>

<%if Request.QueryString("message").count > 0 then DBA_WriteSuccess Request.QueryString("message")%>

<form action="recedit.asp" method="post">
<input type="hidden" name="<%=TableNameKey%>" value="<%=sTable%>">
<input type="hidden" name="<%=PKNameKey%>" value="<%=PrimaryKeys%>">
<input type="hidden" name="<%=ActionNameKey%>" value="<%=action%>">
<input type="hidden" name="<%=KeyNameKey%>" value="<%=Join(key, Separator)%>">
<input type="hidden" name="<%=NextPos%>" id="<%=NextPos%>" value="">
<table border=0 align="center">

<%	i = 0
	for each fld in rec.Fields
		if fld.Type <> adBinary and fld.Type <> adVarBinary and fld.Type <> adLongVarBinary then
			if dba.Tables.Item(sTable).Fields.Item(fld.Name).HasDefault then
				DefaultValue = dba.Tables.Item(sTable).Fields.Item(fld.Name).DefaultValue
			else
				DefaultValue = ""
			end if
%>
	<tr>
		<td valign=top bgcolor="#006699"><font color="white">
			<b><%=fld.Name%></b>&nbsp;(<%=GetTypeName(fld.Type)%>)
			<%if Len(DefaultValue) > 0 then Response.Write "<br>Default:&nbsp;" & DefaultValue%>
		</font></td>
		<td style="border: 1px solid #c6d9ce" valign="top">
<%
			'>>>>>> FORM FOR ALL FIELDS <<<<<<'
			if fld.Type = 203 or fld.Type = 201 then
				Response.Write "<textarea id=""fld" & i & """ name=""" & fld.Name & """ rows=""4"" cols=""40"">"
				if bIsEdit then Response.Write Server.HTMLEncode(CStr(fld.Value))
				Response.Write "</textarea>" & vbCrLf
				if Len(DBA_addTextEditor) > 0 then Response.Write "<input type=""button"" onclick=""javascript:DBA_popupWindow('" & DBA_cfgAddonsFolder & "/" & DBA_addTextEditor & "?fld" & i & "', 'editor', 535, 360)"" class=""button"" value=" & langEdit & ">"
			ElseIf fld.Type = 11 Then
				'this is a boolean value
				Response.Write "<input type=""checkbox"" name=""" & fld.Name & """ value=""1"""
				If bIsEdit and fld.Value = True Then Response.Write " checked "
				Response.Write ">" & vbCrLf
			ElseIf fld.Properties("ISAUTOINCREMENT") then
				Response.Write "AutoNumber (" & fld.Value & ")"
			Else
				Response.Write "<input type=""text"" id=""fld" & i & """ name=""" & fld.Name & """ value="""
				if bIsEdit then Response.Write Replace(CStr(fld.Value), """", "&quot;")
				Response.Write """>" & vbCrLf
				If Len(dba.Tables.Item(sTable).Fields.Item(fld.Name).LookupTable) > 0 Then
					Response.Write	"<input type=""button"" onclick=""javascript:DBA_popupWindow('lookup.asp?id=fld" & i & "&table=" &_
									Server.URLEncode(dba.Tables.Item(sTable).Fields.Item(fld.Name).LookupTable) &_
									"&field=" & Server.URLEncode(dba.Tables.Item(sTable).Fields.Item(fld.Name).LookupField) &_
									"', 'lookup', 640, 400);"" class=""button"" value=""" & langLookup & """>"
				ElseIf Len(DBA_addTextEditor) > 0 and GetTypeName(fld.Type) = "Text" then 
					Response.Write "<input type=""button"" onclick=""javascript:DBA_popupWindow('" & DBA_cfgAddonsFolder & "/" & DBA_addTextEditor & "?fld" & i & "', 'editor', 550, 360)"" class=""button"" value=" & langEdit & ">"
				End If
			End If
%>
		</td>
	</tr>
		<%end if%>
<%		i = i + 1
	Next%>

</table>
<table align=center>
	<tr>
		<td>
			<%if bHasPrev = True then%>
			<input type="submit" value="<%=sName & " + " & langPrev%>" class="button" onclick="javascript:onGoNext('update_prev');">
			<%end if%>
		</td>
		<td>&nbsp;</td>
		<td><input type="submit" value="<%=sName%>" class="button" onclick="javascript:onGoNext('update');"></td>
		<td><input type="reset" value="<%=langReset%>" class="button"></td>
		<td><input type="button" value="<%=langCancel%>" class="button" onclick="javascript:onCancelEdit('<%=sTable%>', '<%=page%>');"></td>
		<td>&nbsp;</td>
		<td>
			<%if bHasNext = True then%>
			<input type="submit" value="<%=sName & " + " & langNext%>" class="button" onclick="javascript:onGoNext('update_next');">
			<%end if%>
		</td>
	</tr>
	<%if action = "edit" then%>
	<tr>
		<td colspan="2" align="right">
			<%if bHasPrev = True then%>
			<input type="submit" value="<< <%=langPrev%>" class="button" onclick="javascript:onGoNext('prev');">
			<%end if%>
		</td>
		<td align=right><input type=submit value=" <%=langFirst%> " class="button" onclick="javascript:onGoNext('first');"></td>
		<td>&nbsp;</td>
		<td><input type="submit" value=" <%=langLast%> " class="button" onclick="javascript:onGoNext('last');"></td>
		<td colspan="2">
			<%if bHasNext = True then%>
			<input type="submit" value="<%=langNext%> >>" class="button" onclick="javascript:onGoNext('next');">
			<%end if%>
		</td>
	</tr>
	<tr><td colspan="7" align=center><hr align="center" width="75%"></td></tr>
	<tr>
		<td colspan="3" align="right"><input type="submit" value=" <%=langAdd%> " class="button" onclick="javascript:onAddNew();"></td>
		<td>&nbsp;</td>
		<td colspan="3"><input type="submit" value="<%=langDelete%>" class="button" onclick="return onDelete();"></td>
	</tr>
	<%end if%>
</table>
</form>
<%
	rec.Close
	set rec = nothing
	call DBA_EndNewTable
	set dba = Nothing
%>
<!--#include file=scripts\inc_footer.inc -->
</body>
</html>
<script language="vbscript" runat="server">
Function GetTypeName(intType)
	Select Case intType
	Case 3			GetTypeName = "Long Integer"
	Case 7			GetTypeName = "Date/Time"
	Case 11			GetTypeName = "Boolean"
	Case 6			GetTypeName = "Currency"
	Case 128,204	GetTypeName = "Binary"
	Case 17			GetTypeName = "Byte"
	Case 131		GetTypeName = "Decimal"
	Case 5			GetTypeName = "Double"
	Case 2			GetTypeName = "Integer"
	Case 4			GetTypeName = "Single"
	Case 72			GetTypeName = "Replication ID"
	Case 203,201	GetTypeName = "Memo"
	Case 202,200	GetTypeName = "Text"
	Case Else		GetTypeName = intType
	End Select
End Function

Function GetPKValues (ByRef pk, ByRef rec, bAsArray)
	dim key, fld
	for each fld in pk
		if Len(fld) > 0 then
			Select Case rec(fld).Type 
				Case adBSTR,adChar,adLongVarChar,adLongVarWChar,adVarChar,adVarWChar,adWChar
					key = key & "'" & Replace(rec(fld), "'", "''") & "'"
				Case adDate,adDBDate, adDBTime, adDBTimeStamp,adFileTime
					key = key & "CDate('" & FormatDateTime(rec(fld), vbLongDate) & " " & FormatDateTime(rec(fld), vbLongTime) & "')"
				Case Else
					key = key & rec(fld)
			End Select
			key = key & Separator
		end if
	Next
	if Right(key, 1) = Separator then key = Left(key, Len(key)-1)
	if bAsArray = True then
		GetPKValues = Split(key, Separator)
	else
		GetPKValues = key
	end if
End Function

Function BuildFilter
	dim filter, field, cmp, criteria, fldType
	filter = ""
	field = Request.QueryString("filter_field").Item
	cmp = Request.QueryString("filter_cmp").Item
	criteria = Request.QueryString("filter_criteria").Item
	
	If Len(field) > 0 and Len(criteria) > 0 then
		fldType = dba.Tables.Item(sTable).Fields.Item(field).GetSQLTypeName()
		If fldType = "TEXT" or fldType = "MEMO" Then
			'remove asterics if only at beginning
			If Left(criteria, 1) = "*" and Right(criteria, 1) <> "*" Then criteria = Mid(criteria, 2)
			criteria = "'" & Replace(criteria, "'", "''") & "'"
		ElseIf fldType = "DATETIME" Then
			criteria = "#" & criteria & "#"
		Else
			If cmp = "LIKE" Then cmp = "="
		End If
		filter = field & " " & cmp & " " & criteria
	End If
	
	BuildFilter = filter
End Function
</script>
