<%@ Language=VBScript %>
<!--#include file=scripts\inc_common.asp -->
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<head>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet" type="text/css">
<title>DBA:Table Data</title>
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
</head>
<body>
		
<%
On Error Resume Next
Dim Separator : Separator = vbTab
dim rec, sTableName, fld, abspage, i, pk, sSQL, key, dba, item
dim pagesize, action, page, PrimaryKeys, sClass
pk = ""
PrimaryKeys = ""
sTableName = Request("table").Item
action = CStr(Request("action").Item)

pagesize = 10
page = 1
if IsNumeric(Request("pagesize").Item) then pagesize = CInt(Request("pagesize").Item)
if pagesize < 1 then pagesize = 10
if IsNumeric(Request("page").Item) then page = CInt(Request("page").Item)
if pagesize < 1 then pagesize = 10
if page < 1 then page = 1

call DBA_WriteNavigation

set dba = new DBAdmin
dba.Connect Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword)
If not dba.IsOpen then Response.Redirect "database.asp"
DBA_BeginNewTable sTableName & langDataForTable, "", "90%"
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
for each item in dba.Tables.Item(sTableName).Indexes.Items
	if item.IsPrimary then PrimaryKeys = PrimaryKeys & item.TargetField & Separator
next
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
			<%=langPageSize%>&nbsp;
			<select name="pagesize">
				<option value="5">5</option>
				<option value="10">10</option>
				<option value="15">15</option>
				<option value="25">25</option>
				<option value="50">50</option>
			</select>
			<input type=hidden name="table" value="<%=sTableName%>">
			<input type=submit value="<%=langSubmit%>" class="button">
		</p>
	</form>
</div>
	<p align="left">
	<%if abspage > 1 then%>
		<a href="data.asp?table=<%=Server.URLEncode(sTableName)%>&amp;page=<%=(abspage - 1)%>&amp;pagesize=<%=pagesize%>&order=<%=Server.URLEncode(Request.QueryString("order"))%>"><font size="1">&laquo;&nbsp;<%=langPrev%></font></a>
	<%end if%>
	<%for i=1 to rec.PageCount
		if i = abspage then%>
			<font size="2">[<%=i%>]</font>&nbsp;
	<%	else%>
			<font size="1">&nbsp;[<a href="data.asp?table=<%=Server.URLEncode(sTableName)%>&amp;page=<%=i%>&amp;pagesize=<%=pagesize%>&order=<%=Server.URLEncode(Request.QueryString("order"))%>"><%=i%></a>]&nbsp;</font>
	<%	end if
	Next
	if abspage < rec.PageCount and abspage > 0 then%>
		<a href="data.asp?table=<%=Server.URLEncode(sTableName)%>&amp;page=<%=(abspage + 1)%>&amp;pagesize=<%=pagesize%>&order=<%=Server.URLEncode(Request.QueryString("order"))%>"><font size="1"><%=langNext%>&nbsp;&raquo;</font></a>
	<%end if
	i = 0
	%>
	</p>
<table align="center" border="1" class="DataTable">
<tr>
	<th>*</th>
<%for each fld in rec.Fields%>
	<th>
		<%if dba.Tables.Item(sTableName).Fields.Item(fld.Name).IsPrimaryKey then%>
			<img src="images/key.gif" border="0" WIDTH="16" HEIGHT="16">
		<%end if%>
		<A href="data.asp?table=<%=Server.URLEncode(sTableName)%>&order=<%=Server.URLEncode(fld.Name & " ASC")%>" title="<%=langSortAscending%>"><font color=white><%=fld.Name%></font></A>&nbsp;<A href="data.asp?table=<%=Server.URLEncode(sTableName)%>&order=<%=Server.URLEncode(fld.Name & " DESC")%>" title="<%=langSortDescending%>"><font color=white>&darr;</font></A>
	</th>
<%next%>
</tr>

<%
	do while not rec.EOF and i < rec.PageSize and Err = 0
		if sClass = "oddrow" then sClass = "evenrow" else sClass = "oddrow"
%>
<tr class="<%=sClass%>" onmouseover="style.backgroundColor='#ffdfbf'" onmouseout="style.backgroundColor=''">
	<td valign="top">
	<%if Len(PrimaryKeys) > 0 then
		sSQL = ""
		for each fld in dba.Tables.Item(sTableName).Fields.Items
			if fld.IsPrimaryKey then
				Select Case fld.FieldType 
					Case adBSTR,adChar,adLongVarChar,adLongVarWChar,adVarChar,adVarWChar,adWChar
						sSQL = sSQL & "'" & Replace(rec(fld.Name), "'", "''") & "'"
					Case adDate,adDBDate, adDBTime, adDBTimeStamp,adFileTime
						sSQL = sSQL & "CDate('" & DBA_FormatDateTime(rec(fld.Name).Value) & "')"
					Case Else
						sSQL = sSQL & rec(fld.Name)
				End Select
				sSQL = sSQL & Separator
			end if
		Next
	%>
		<a href="recedit.asp?action=edit&amp;pk=<%=Server.URLEncode(PrimaryKeys)%>&amp;key=<%=Server.URLEncode(sSQL)%>&amp;table=<%=Server.URLEncode(sTableName)%>&amp;page=<%=page%>&amp;pagesize=<%=pagesize%>"><img src="images/edit.gif" alt="<%=langEditRecord%>" border="0" WIDTH="16" HEIGHT="16"></a>
		<a href="javascript:deleteRecord('<%=Server.URLEncode(Replace(PrimaryKeys, "'", "\'"))%>','<%=Server.URLEncode(Replace(sSQL, "'", "\'"))%>')"><img src="images/delete.gif" alt="<%=langDeleteRecord%>" border="0" WIDTH="16" HEIGHT="16"></a>
	<%end if%>
	</td>
	<%for each fld in rec.Fields%>
		<td valign="top" align="center" class="DataTD">
		<%if fld.Type <> adBinary then%>
			<%if Len(fld.Value) > 0 then
				Response.Write Server.HTMLEncode(fld.Value)
			else
				Response.Write "&nbsp;"
			end if
		else
			Response.Write "&lt;" & langBinaryData & "&gt;"
		end if%>
		</td>
	<%next%>
</tr>
<%	rec.MoveNext
	i = i + 1 
loop%>

</table>		

	<p align="left">
	<%if abspage > 1 then%>
		<a href="data.asp?table=<%=Server.URLEncode(sTableName)%>&amp;page=<%=(abspage - 1)%>&amp;pagesize=<%=pagesize%>&order=<%=Server.URLEncode(Request.QueryString("order"))%>"><font size="1">&laquo;&nbsp;<%=langPrev%></font></a>
	<%end if%>
	<%for i=1 to rec.PageCount
		if i = abspage then%>
			<font size="2">[<%=i%>]</font>&nbsp;
	<%	else%>
			<font size="1">&nbsp;[<a href="data.asp?table=<%=Server.URLEncode(sTableName)%>&amp;page=<%=i%>&amp;pagesize=<%=pagesize%>&order=<%=Server.URLEncode(Request.QueryString("order"))%>"><%=i%></a>]&nbsp;</font>
	<%	end if
	Next
	if abspage < rec.PageCount and abspage > 0 then%>
		<a href="data.asp?table=<%=Server.URLEncode(sTableName)%>&amp;page=<%=(abspage + 1)%>&amp;pagesize=<%=pagesize%>&order=<%=Server.URLEncode(Request.QueryString("order"))%>"><font size="1"><%=langNext%>&nbsp;&raquo;</font></a>
	<%end if%>
	</p>

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
