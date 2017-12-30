<%@ Language=VBScript %>
<!--#include file=scripts\inc_common.asp -->
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<head>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet" type="text/css">
<title>DBA:Tables list</title>
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
</head>
<body>

<%	call DBA_WriteNavigation%>

<%
	On Error Resume Next
	dim dba, key, sClass, dic, action
	set dba = new DBAdmin
	dba.Connect Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword)
	action = CStr(Request("action").Item)
	
	DBA_BeginNewTable langTablesList, "", "90%"
	if dba.HasError then DBA_WriteError dba.LastError
	Select Case action
		Case "create"
			dba.CreateTable Request.Form("tablename").Item
			if dba.HasError then DBA_WriteError dba.LastError
		Case "delete"
			dba.DeleteTable Request.QueryString("table").Item
	End Select
%>

<table align="center" border="0" cellpadding="2" cellspacing="1" width="90%">
	<tr>
		<th><%=langTableName%></th>
		<th><%=langCreated%></th>
		<th><%=langModified%></th>
		<th><%=langActions%></th>
	</tr>
<%	set dic = dba.Tables
	for each key in dic.Keys
		if sClass = "oddrow" then sClass = "evenrow" else sClass = "oddrow"
%>
	<tr class="<%=sClass%>" onmouseover="style.backgroundColor='#ffdfbf'" onmouseout="style.backgroundColor=''">
		<td><%=dic.Item(key).Name%></td>
		<td align="right"><%=dic.Item(key).DateCreated%></td>
		<td align="right"><%=dic.Item(key).DateModified%></td>
		<td align="right">
			<a href="structure.asp?table=<%=Server.URLEncode(dic.Item(key).Name)%>"><img src="images/structure.gif" alt="<%=langViewTableStructureAlt%>" border="0" width="16" height="16"></a>
			&nbsp;<a href="data.asp?table=<%=Server.URLEncode(dic.Item(key).Name)%>"><img src="images/table.gif" alt="<%=langViewTableDataAlt%>" border="0" width="16" height="16"></a>
			&nbsp;<a href="recedit.asp?action=edit&amp;table=<%=Server.URLEncode(dic.Item(key).Name)%>"><img src="images/cycle.gif" alt="<%=langTableNavigateAlt%>" border="0" width="16" height="16"></a>
			&nbsp;<a href="tablelist.asp?action=delete&amp;table=<%=Server.URLEncode(dic.Item(key).Name)%>" onclick="return confirm('<%=Replace(langSureToDeleteTable, "$table_name", dic.Item(key).Name)%>?');"><img src="images/delete.gif" alt="<%=langDeleteTableAlt%>" border="0" width="16" height="16"></a>
		</td>
	</tr>
<%	next%>
</table>

<form action="tablelist.asp" method="post">
<input type="hidden" name="action" value="create">
<table align="center" border="0">
	<tr>
		<td><%=langNewTableName%></td>
		<td><input type="text" name="tablename"></td>
		<td><input type="submit" name="submit" value="<%=langCreateNewTable%>" class="button"></td>
	</tr>
</table>
</form>
<%
	call DBA_EndNewTable
	set dba = Nothing
%>
<!--#include file=scripts\inc_footer.inc -->
</body>
</html>
