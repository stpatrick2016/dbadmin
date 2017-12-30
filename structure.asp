<%@ Language=VBScript %>
<!--#include file=scripts\inc_common.asp -->
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<head>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet">
<title>DBA:Table Structure</title>
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
<script language="javascript" type="text/javascript">
//<!--
function deleteColumn(name){
	var str = '<%=Replace(langAreYouSureToDelete, """", "\""")%>';
	str = str.replace("\$name", name);
	if(confirm(str)){
		document.location.replace('structure.asp?table=<%=Server.URLEncode(Request("table"))%>&field=' + name + '&action=delete_field');
	}
}
function onFieldTypeChange(newType){
	var isText = newType == "TEXT" || newType == "MEMO" ? true : false;
	document.getElementById('trZeroLength').style.display = isText ? "" : "none";
	document.getElementById('trCompress').style.display = isText ? "" : "none";
}
//-->
</script>
</head>
<body>
<%	call DBA_WriteNavigation%>

<%
	On Error Resume Next
	dim dba, dic, strTable, item, action, sClass, sPrimaryIndexName, oTable
	strTable = CStr(Request("table"))
	action = CStr(Request("action"))
	set dba = new DBAdmin
	dba.Connect Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword)
	
	DBA_BeginNewTable strTable & langTableStructure, "", "90%"
	If not dba.IsOpen then Response.Redirect "database.asp"
	if dba.HasError then DBA_WriteError dba.LastError
	
	set oTable = dba.Tables.Item(strTable)
	'perform requested actions first
	Select Case action
		Case "create_field"
			set item = new DBAField
			item.Name = Request.Form("fname").Item
			item.FieldType = Request.Form("type").Item
			item.DefaultValue = Request.Form("default").Item
			item.IsNullable = Request.Form("nullable").Item
			item.MaxLength = Request.Form("length").Item
			item.Description = Request.Form("description").Item
			item.AllowZeroLength = Request.Form("zero_length").Item
			item.Compressed = Request.Form("compress").Item
			oTable.CreateField item, Request.Form("indexed").Item
			set item = nothing
			if dba.HasError then DBA_WriteError dba.LastError
		Case "delete_field"
			oTable.DeleteField Request.QueryString("field").Item
			if dba.HasError then DBA_WriteError dba.LastError
		Case "key"
			oTable.CreateIndex "", Request.QueryString("field").Item, True, True
			if dba.HasError then DBA_WriteError dba.LastError
		Case "delete_index"
			oTable.DeleteIndex Request.QueryString("index").Item, Request.QueryString("field").Item
			if dba.HasError then DBA_WriteError dba.LastError
		Case "create_index"
			oTable.CreateIndex Request.Form("index").Item, Request.Form("field").Item, Request.Form("unique").Item, False
			if dba.HasError then DBA_WriteError dba.LastError
		Case "update_field"
			set item = oTable.Fields.Item(Request.Form("oldname").Item)
			item.Name = Request.Form("fname").Item
			item.FieldType = Request.Form("type").Item
			item.MaxLength = Request.Form("length").Item
			item.DefaultValue = Request.Form("default").Item
			item.Description = Request.Form("description").Item
			item.AllowZeroLength = Request.Form("zero_length").Item
			item.IsNullable = Request.Form("nullable").Item
			item.UpdateBatch
			if dba.HasError then DBA_WriteError dba.LastError
	End Select
	
	'find out primary index name
	sPrimaryIndexName = ""
	set dic = dba.Tables.Item(strTable).Indexes
	for each item in dic.Items
		if item.IsPrimary then
			sPrimaryIndexName = item.Name
			Exit For
		end if
	next
	
	set dic = dba.Tables.Item(strTable).Fields
%>

<!--FIELDS-->
<table align="center" width="90%" border="0" cellpadding="3" cellspacing="1">
	<tr>
		<th><%=langOrdinal%></th>
		<th><%=langName%></th>
		<th><%=langDataType%></th>
		<th><%=langNullable%></th>
		<th><%=langMaxLength%></th>
		<th><%=langDefaultValue2%></th>
		<th><%=langDescription%></th>
		<th><%=langActions%></th>
	</tr>
<%	
	for each item in dic.Items
		WriteFieldTR item
	next
%>
</table>
<%
	call DBA_EndNewTable
	if action = "edit" then 
		call WriteFieldEditForm(dic.Item(Request.QueryString("field").Item))
	else
		set item = new DBAField
		item.Ordinal = dic.Count + 1
		call WriteFieldEditForm(item)
		set item = nothing
	end if
%>

<!--INDEXES-->
<%
	DBA_BeginNewTable strTable & langTableIndexes, "", "90%"
%>
<form action="structure.asp" method="post">
<table align="center" width="90%" border="0" cellpadding="3" cellspacing="1">
	<tr>
		<th><%=langIndexName%></th>
		<th><%=langColumn%></th>
		<th><%=langUnique%></th>
		<th><%=langActions%></th>
	</tr>
<%
	set dic = dba.Tables.Item(strTable).Indexes
	sClass = ""
	for each item in dic.Items
		if sClass = "oddrow" then sClass = "evenrow" else sClass = "oddrow"
%>
	<tr class="<%=sClass%>" onmouseover="style.backgroundColor='#ffdfbf'" onmouseout="style.backgroundColor=''">
		<td>
<%			if item.IsPrimary then%>
				<img src="images/key.gif" border="0" width="16" height="16" alt="<%=langPrimaryColumnAlt%>">
<%			end if%>
		<%=item.Name%>
		</td>
		<td><%=item.TargetField%></td>
		<td align="center">
			<%if item.IsUnique then%>
			<img src="images/check.gif" border="0" width="16" height="16" alt="<%=langUniqueIndexAlt%>">
			<%end if%>
			&nbsp;
		</td>
		<td align="center">
			<a href="structure.asp?table=<%=Server.URLEncode(strTable)%>&amp;action=delete_index&amp;index=<%=Server.URLEncode(item.Name)%>&field=<%=Server.URLEncode(item.TargetField)%>"><img src="images/delete.gif" alt="<%=langDeleteIndexAlt%>" border="0" width="16" height="16"></a>
		</td>
	</tr>
<%	next%>
	<tr>
		<th align="left"><input type="text" name="index" size="10"></th>
		<th align="left">
			<select name="field">
<%	for each item in dba.Tables.Item(strTable).Fields.Items%>
				<option value="<%=item.Name%>"><%=item.Name%></option>
<%	next%>
			</select>
		</th>
		<th align="center"><input type="checkbox" name="unique" value="1"></th>
		<th align="right"><input type="submit" name="submit" value="<%=langCreateIndex%>" class="button"></th>
	</TR>
</table>
<input type="hidden" name="table" value="<%=strTable%>">
<input type="hidden" name="action" value="create_index">
</form>

<!--GENERATED SQL QUERY-->
<%
	call DBA_EndNewTable
	DBA_BeginNewTable strTable & langCreateTableQuery, langCreateTableQueryAlt, "75%"
%>
	<div id="divSQL" align="left"><%=HighlightSQL(dba.Tables.Item(strTable).SQL)%></div>
<%	if InStr(1, Request.ServerVariables("HTTP_USER_AGENT"), "MSIE") > 0 then%>
		<p>&nbsp;</p>
		<div align="center"><input type="button" value="<%=langCopyToClipboard%>" onclick="javascript:copyToClipboard(document.getElementById('divSQL'));" class="button"></div>
<%	end if%>
<%
	call DBA_EndNewTable
	set dba = Nothing
%>
<!--#include file=scripts\inc_footer.inc -->
</body>
</html>

<%
	Sub WriteFieldTR(byref fld)
		if sClass = "oddrow" then sClass = "evenrow" else sClass = "oddrow"
%>
		<tr class="<%=sClass%>" onmouseover="style.backgroundColor='#ffdfbf'" onmouseout="style.backgroundColor=''">
			<td align=center><%=fld.Ordinal%></td>
			<td>
				<%if fld.IsPrimaryKey then%>
					<img src="images/key.gif" border="0" width="16" height="16" alt="<%=langPrimaryColumnAlt%>">
				<%end if%>
				<%=fld.Name%>
			</td>
			<td><%=fld.GetTypeName%></td>
			<td align="center">
				<%if fld.IsNullable then%>
				<img src="images/check.gif" border="0" width="16" height="16">
				<%end if%>
				&nbsp;
			</td>
			<td><%if fld.MaxLength > 0 then Response.Write fld.MaxLength else Response.Write "&nbsp;"%></td>
			<td><%=fld.DefaultValue%>&nbsp;</td>
			<td><%=fld.Description%>&nbsp;</td>
			<td align="center">
				<a href="structure.asp?table=<%=Server.URLEncode(strTable)%>&amp;field=<%=Server.URLEncode(fld.Name)%>&amp;action=edit"><img src="images/edit.gif" alt="<%=langEditField%>" border="0" width="16" height="16"></a>&nbsp;
				<%if fld.IsPrimaryKey then%>
					<a href="structure.asp?table=<%=Server.URLEncode(strTable)%>&amp;field=<%=Server.URLEncode(fld.Name)%>&index=<%=Server.URLEncode(sPrimaryIndexName)%>&amp;action=delete_index"><img src="images/un_key.gif" alt="<%=langRemovePK%>" border="0" width="16" height="16"></a>&nbsp;
				<%else%>
					<a href="structure.asp?table=<%=Server.URLEncode(strTable)%>&amp;field=<%=Server.URLEncode(fld.Name)%>&amp;action=key"><img src="images/key.gif" alt="<%=langSetAsPK%>" border="0" width="16" height="16"></a>&nbsp;
				<%end if%>
				<a href="javascript:deleteColumn('<%=Server.URLEncode(fld.Name)%>')"><img src="images/delete.gif" alt="<%=langDeleteField%>" border="0" width="16" height="16"></a>
			</td>
		</tr>
<%	End Sub%>

<%	
	Sub WriteFieldEditForm(byref fld)
		dim isText, isNewField, strNullable, strCompress
		if Len(fld.Name) =  0 then 
			DBA_BeginNewTable strTable & "&nbsp;:&nbsp;" & langAddNewColumn, "", "50%" 
			action = "create_field"
			isNewField = True
			strNullable = "<input type=""checkbox"" name=""nullable"" value=""-1"">"
			strCompress = "<input type=""checkbox"" name=""compress"" value=""-1"">"
		else 
			DBA_BeginNewTable strTable & "&nbsp;:&nbsp;" & langEditField, "", "50%"
			action = "update_field"
			isNewField = False
			strNullable = "<input type=""hidden"" name=""nullable"" value=""" & CInt(fld.IsNullable) & """>" & CStr(fld.IsNullable)
			strCompress = CStr(fld.Compressed)
		end if
		if fld.GetSQLTypeName = "TEXT" or fld.getSQLTypeName = "MEMO" then isText = True else IsText = False
%>
<!--Field editing form -->
		<form action="structure.asp" method="post">
		<input type="hidden" name="oldname" value="<%=fld.Name%>">
		<input type="hidden" name="table" value="<%=strTable%>">
		<input type="hidden" name="action" value="<%=action%>">
		<table align="center" border="0" cellpadding="3" cellspacing="1">
			<tr>
				<td><%=langOrdinal%></td>
				<td><%=fld.Ordinal%></td>
			</tr>
			<tr>
				<td><%=langName%></td>
				<td><input type="text" name="fname" value="<%=fld.Name%>"></td>
			</tr>
			<tr>
				<td><%=langDataType%></td>
				<td><select name="type" onchange="javascript:onFieldTypeChange(this.options[this.selectedIndex].value);">
					<option value="<%=fld.GetSQLTypeName%>"><%=fld.GetTypeName%></option>
					<option value="DATETIME">Date/Time</option>
					<option value="LONG">Long Integer</option>
					<option value="TEXT">Text</option>
					<option value="COUNTER">AutoNumber</option>
					<option value="MEMO">Memo</option>
					<option value="MONEY">Currency</option>
					<option value="BINARY">Binary</option>
					<option value="TINYINT">Byte</option>
					<option value="DECIMAL">Decimal</option>
					<option value="FLOAT">Double</option>
					<option value="INTEGER">Integer</option>
					<option value="REAL">Single</option>
					<option value="BIT">Boolean</option>
					<option value="UNIQUEIDENTIFIER">Replication ID</option>
				</select></td>
			</tr>
			<tr>
				<td><%=langNullable%></td>
				<td><%=strNullable%></td>
			</tr>
			<tr>
				<td><%=langMaxLength%></td>
				<td><input type="text" name="length" size="5" value="<%=fld.MaxLength%>"></td>
			</tr>
			<tr>
				<td><%=langDefaultValue2%></td>
				<td><input type="text" name="default" size="20" value="<%=fld.DefaultValue%>"></td>
			</tr>
			<tr>
				<td><%=langDescription%></td>
				<td><input type="text" name="description" value="<%=fld.Description%>"></td>
			</tr>
			<tr id="trZeroLength" style="<%if not isText then Response.Write "display:none"%>">
				<td><%=langAllowZeroLength%></td>
				<td><input type="checkbox" name="zero_length" value="-1" <%if fld.AllowZeroLength then Response.Write " checked "%>></td>
			</tr>
			<tr id="trCompress" style="<%if not isText then Response.Write "display:none"%>">
				<td><%=langUnicodeCompress%></td>
				<td><%=strCompress%></td>
			</tr>
<%		if isNewField then%>
			<tr>
				<td><%=langIndexed%></td>
				<td><select name="indexed">
					<option value="0"><%=langNo%></option>
					<option value="1"><%=langIndxedDuplicates%></option>
					<option value="2"><%=langIndexedUnique%></option>
				</select></td>
			</tr>
<%		end if%>
			<tr>
				<td align="center" colspan="2">
					<input type="submit" name="submit" value="<%=langUpdate%>" class="button">
				<%if Len(fld.Name) > 0 then%>
					<input type="button" value="<%=langCancel%>" class="button" onclick="javascript:window.location.replace('structure.asp?table=<%=Server.URLEncode(strTable)%>');">
				<%end if%>
				</td>
			</tr>
		</table>
		</form>
		<!--End of form-->
<%	
		DBA_EndNewTable
	End Sub
%>
