<%@ Language=VBScript %>
<!--#include file=scripts\inc_common.asp -->
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<head>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet" type="text/css">
<title>DBA:<%=langCaptionData%></title>
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
</head>
<body>
		
<%
On Error Resume Next
Dim Separator : Separator = vbTab
dim rec, sTableName, fld, abspage, i, pk, sSQL, key, dba, item
dim pagesize, action, page, PrimaryKeys, sClass, arrTemp, strFilter
pk = ""
PrimaryKeys = ""
sTableName = Request("table").Item
action = CStr(Request("action").Item)

pagesize = 10
page = 1
if IsNumeric(Request("pagesize").Item) then pagesize = CInt(Request("pagesize").Item)
if IsNumeric(Request("page").Item) then page = CInt(Request("page").Item)
if pagesize < 1 then pagesize = StpProfile.GetProfileNumber("settings", "page_size", 10)
if page < 1 then page = 1

call DBA_WriteNavigation

set dba = new DBAdmin
dba.Connect Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword)
If not dba.IsOpen then Response.Redirect "database.asp"
DBA_BeginNewTable sTableName & langDataForTable, "", "90%", ""
if dba.HasError then DBA_WriteError dba.LastError

'delete record
if Request.QueryString("action") = "delete" then
	sSQL = "DELETE FROM [" & sTableName & "] WHERE"
	fld = ""
	pk = Split(Request.QueryString("pk"), Separator)
	key = Split(Request.QueryString("key"), Separator)
	for i=0 to UBound(pk)
		if Len(pk(i)) > 0 and Len(key(i)) > 0 then
			sSQL = sSQL & fld & " [" & pk(i) & "]=" & key(i)
			fld = " AND"
		end if
	Next
	call dba.RunScript(sSQL, False, False, null)
	if dba.HasError then DBA_WriteError dba.LastError
end if


sSQL = "SELECT * FROM [" & sTableName & "]"
if Len(Request.QueryString("order")) > 0 then sSQL = sSQL & " ORDER BY " & Request.QueryString("order")
set rec = dba.RunScript(sSQL, False, False, null)
rec.CacheSize = pagesize
rec.PageSize = pagesize
if dba.HasError then DBA_WriteError dba.LastError

if rec.PageCount > 0 then rec.AbsolutePage = page
abspage = rec.AbsolutePage

'retrieve string with primary keys names
If dba.Tables.Exists(sTableName) Then
	for each item in dba.Tables.Item(sTableName).Indexes.Items
		if item.IsPrimary then PrimaryKeys = PrimaryKeys & item.TargetField & Separator
	next
End If
if Right(PrimaryKeys, 1) = Separator then PrimaryKeys = Left(PrimaryKeys, Len(PrimaryKeys)-1)
%>
<div style="border: 1px #006699 solid; padding-left: 15px">
<p align="left">
<%if Len(PrimaryKeys) > 0 then%>
*&nbsp;<img src="images/add.gif" border="0" WIDTH="16" HEIGHT="16"><a href="recedit.asp?table=<%=Server.URLEncode(sTableName)%>&amp;pk=<%=Server.URLEncode(PrimaryKeys)%>&amp;page=<%=page%>&amp;pagesize=<%=pagesize%>"><%=langAddRecord%></a>&nbsp;
<%end if%>
*&nbsp;<img src="images/refresh.gif" border="0" WIDTH="16" HEIGHT="16"><a href="data.asp?table=<%=Server.URLEncode(sTableName)%>"><%=langRefreshTable%></a>&nbsp;
*&nbsp;<img src="images/xml.gif" border="0" WIDTH="16" HEIGHT="16"><a href="export_xml.asp?sql=<%=Server.URLEncode(sSQL)%>" alt="<%=langXMLExportAlt%>"><%=langXMLExport%></a>&nbsp;
*&nbsp;<img src="images/excel.gif" border="0" WIDTH="16" HEIGHT="16"><a href="export_csv.asp?sql=<%=Server.URLEncode(sSQL)%>" alt="<%=langExcelExportAlt%>"><%=langExcelExport%></a>&nbsp;*
</p>
<%if Len(PrimaryKeys) = 0 then DBA_WriteError langNoPrimaryKey%>
	<form action="data.asp" method="get">
		<p align="left">
			<%=langFilter%>&nbsp;
			<select name="filter_field">
				<option value=""></option>
<%	For Each item In dba.Tables.Item(sTableName).Fields.Items%>
				<option value="<%=item.Name%>"><%=item.Name%></option>
<%	Next%>
			</select>
			<select name="filter_cmp">
				<option value="=">=</option>
				<option value=">">></option>
				<option value="<"><</option>
				<option value=">=">>=</option>
				<option value="<="><=</option>
				<option value="<>"><></option>
				<option value="LIKE">LIKE</option>
			</select>
			<input type="text" name="filter_criteria" size="10">
			
			&nbsp;&nbsp;
			<%=langPageSize%>&nbsp;
			<select name="pagesize">
				<%=DBA_GetComboOptions(5, 50, 5, pagesize)%>
			</select>
			<input type=hidden name="table" value="<%=sTableName%>">
			<input type=submit value="<%=langSubmit%>" class="button">
		</p>
	</form>
</div>
<%=getPagingControl(rec.RecordCount, abspage, pagesize, "&amp;table=" & Server.URLEncode(sTableName) & "&amp;order=" & Server.URLEncode(Request.QueryString("order")))%>
<table align="center" border="1" class="DataTable">
<tr>
	<th>*</th>
<%for each fld in rec.Fields%>
	<th>
		<%if dba.Tables.Item(sTableName).Fields.Item(fld.Name).IsPrimaryKey then%>
			<img src="images/key.gif" border="0" WIDTH="16" HEIGHT="16">
		<%end if%>
		<A href="data.asp?table=<%=Server.URLEncode(sTableName)%>&order=<%=Server.URLEncode("[" & fld.Name & "] ASC")%>" title="<%=langSortAscending%>"><font color=white><%=fld.Name%></font></A>&nbsp;<A href="data.asp?table=<%=Server.URLEncode(sTableName)%>&order=<%=Server.URLEncode("[" & fld.Name & "] DESC")%>" title="<%=langSortDescending%>"><font color=white>&darr;</font></A>
	</th>
<%next%>
</tr>

<%
	if rec.State <> adStateClosed then
		strFilter = BuildFilter()
		if Len(strFilter) > 0 then rec.Filter = strFilter
		do while not rec.EOF and i < rec.PageSize
			if sClass = "oddrow" then sClass = "evenrow" else sClass = "oddrow"
%>
<tr class="<%=sClass%>" onmouseover="style.backgroundColor='#ffdfbf'" onmouseout="style.backgroundColor=''">
	<td valign="top">
	<%	if Len(PrimaryKeys) > 0 then
			sSQL = ""
			arrTemp = Split(PrimaryKeys, Separator)
			for each item in arrTemp
				set fld = dba.Tables.Item(sTableName).Fields.Item(item)
				Select Case fld.FieldType 
					Case adBSTR,adChar,adLongVarChar,adLongVarWChar,adVarChar,adVarWChar,adWChar
						sSQL = sSQL & "'" & Replace(rec(fld.Name), "'", "''") & "'"
					Case adDate,adDBDate, adDBTime, adDBTimeStamp,adFileTime
						sSQL = sSQL & "CDate('" & DBA_FormatDateTime(rec(fld.Name).Value) & "')"
					Case Else
						sSQL = sSQL & rec(fld.Name)
				End Select
				sSQL = sSQL & Separator
			Next
	%>
		<a href="recedit.asp?action=edit&amp;pk=<%=Server.URLEncode(PrimaryKeys)%>&amp;key=<%=Server.URLEncode(sSQL)%>&amp;table=<%=Server.URLEncode(sTableName)%>&amp;page=<%=page%>&amp;pagesize=<%=pagesize%>"><img src="images/edit.gif" alt="<%=langEditRecord%>" border="0" WIDTH="16" HEIGHT="16"></a>
		<a href="javascript:deleteRecord('<%=Server.URLEncode(Replace(PrimaryKeys, "'", "\'"))%>','<%=Server.URLEncode(Replace(sSQL, "'", "\'"))%>')"><img src="images/delete.gif" alt="<%=langDeleteRecord%>" border="0" WIDTH="16" HEIGHT="16"></a>
	<%	end if%>
	</td>
	<%	for each fld in rec.Fields%>
		<td valign="top" align="center" class="DataTD">
		<%	if fld.Type <> adBinary and fld.Type <> adVarBinary then%>
			<%	if Len(fld.Value) > 0 then
					Response.Write Server.HTMLEncode(fld.Value)
				else
					Response.Write "&nbsp;"
				end if
			else
				Response.Write "&lt;" & langBinaryData & "&gt;"
			end if%>
		</td>
	<%	next%>
</tr>
<%		rec.MoveNext
		i = i + 1 
		loop
	end if%>

</table>		

<%=getPagingControl(rec.RecordCount, abspage, pagesize, "&amp;table=" & Server.URLEncode(sTableName) & "&amp;order=" & Server.URLEncode(Request.QueryString("order")))%>

<%
rec.Close
call DBA_EndNewTable
set dba = Nothing
%>
<!--#include file=scripts\inc_footer.inc -->
</body>
<script LANGUAGE="javascript">
<!--
function deleteRecord(pk,key){
	if(confirm("<%=langSureToDeleteRecord%> " + key + "?")){
		document.location.replace("data.asp?action=delete&key=" + escape(key) + "&pk=" + escape(pk) + "&table=<%=sTableName%>&page=<%=page%>&pagesize=<%=pagesize%>");
	}
}
//-->
</script>
</html>

<%
	Function BuildFilter
		dim filter, field, cmp, criteria, fldType
		filter = ""
		field = Request.QueryString("filter_field").Item
		cmp = Request.QueryString("filter_cmp").Item
		criteria = Request.QueryString("filter_criteria").Item
		
		If Len(field) > 0 and Len(criteria) > 0 then
			fldType = dba.Tables.Item(sTableName).Fields.Item(field).GetSQLTypeName()
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
%>