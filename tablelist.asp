<%@ Language=VBScript %>
<!--#include file=scripts\inc_common.asp -->
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<head>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet" type="text/css">
<title>DBA:<%=langCaptionTablesList%></title>
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
</head>
<body>

<%	call DBA_WriteNavigation%>

<%
	if not DBAE_DEBUG Then On Error Resume Next
	dim dba, sClass, action, tbl, tableColor, tableImage
	dim showSysTables
	set dba = new DBAdmin
	dba.Connect Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword)
	action = CStr(Request("action").Item)
	showSysTables = CBool(StpProfile.GetProfileNumber("settings", "sys_tables", 0))
	
	DBA_BeginNewTable langTablesList, "", "90%", ""
	if dba.HasError then DBA_WriteError dba.LastError
	Select Case action
		Case "create"
			dba.CreateTable Request.Form("tablename").Item
			if dba.HasError then DBA_WriteError dba.LastError
		Case "delete"
			dba.DeleteTable Request.QueryString("table").Item
			if dba.HasError then DBA_WriteError dba.LastError
		Case "do_rename"
			dba.Tables.Item(Request.Form("table").Item).Name = Request.Form("tablename").Item
			if dba.HasError then DBA_WriteError dba.LastError
	End Select
%>

<table align="center" border="0" cellpadding="2" cellspacing="1" width="90%">
	<tr>
		<th><%=langTableName%></th>
		<th><%=langCreated%></th>
		<th><%=langModified%></th>
		<th><%=langActions%></th>
	</tr>
<%	
	for each tbl in dba.Tables.Items
		If not tbl.IsSystem or showSysTables Then
			Select Case tbl.TableType
				Case "SYSTEM TABLE", "ACCESS TABLE"
					tableColor = "#808080"
					tableImage = ""
				Case "LINK", "ALIAS"
					tableColor = "#008000"
					tableImage = "<img src=""images/linked.gif"" border=""0"">&nbsp;"
				Case Else
					tableColor = ""
					tableImage = ""
			End Select
			if sClass = "oddrow" then sClass = "evenrow" else sClass = "oddrow"
%>
	<tr class="<%=sClass%>" onmouseover="style.backgroundColor='#ffdfbf'" onmouseout="style.backgroundColor=''">
		<td><%=tableImage%><font color="<%=tableColor%>"><%=tbl.Name%></font></td>
		<td align="right"><%=tbl.DateCreated%></td>
		<td align="right"><%=tbl.DateModified%></td>
		<td align="right">
			<a href="structure.asp?table=<%=Server.URLEncode(tbl.Name)%>"><img src="images/structure.gif" alt="<%=langViewTableStructureAlt%>" border="0" width="16" height="16"></a>
			&nbsp;<a href="data.asp?table=<%=Server.URLEncode(tbl.Name)%>"><img src="images/table.gif" alt="<%=langViewTableDataAlt%>" border="0" width="16" height="16"></a>
			&nbsp;<a href="recedit.asp?action=edit&amp;table=<%=Server.URLEncode(tbl.Name)%>"><img src="images/cycle.gif" alt="<%=langTableNavigateAlt%>" border="0" width="16" height="16"></a>
<%			If StpProfile.ComponentAvailable("ADOX") and not tbl.IsLinked and not tbl.IsSystem Then%>
				&nbsp;<a href="tablelist.asp?action=rename&amp;table=<%=Server.URLEncode(tbl.Name)%>"><img src="images/rename.gif" alt="<%=langRenameTableAlt%>" border="0" width="16" height="16"></a>
<%			End If%>
<%			If not tbl.IsSystem Then%>
				&nbsp;<a href="tablelist.asp?action=delete&amp;table=<%=Server.URLEncode(tbl.Name)%>" onclick="return confirm('<%=Replace(Replace(langSureToDeleteTable, "$table_name", tbl.Name), "'", "\'")%>?');"><img src="images/delete.gif" alt="<%=langDeleteTableAlt%>" border="0" width="16" height="16"></a>
<%			End If%>
		</td>
	</tr>
<%	
		End If
	next
%>
</table>

<form action="tablelist.asp" method="post">
<%	if action = "rename" Then%>
<input type="hidden" name="table" value="<%=Request.QueryString("table").Item%>">
<input type="hidden" name="action" value="do_rename">
<%	else%>
<input type="hidden" name="action" value="create">
<%	end if%>
<table align="center" border="0">
<%	if action = "rename" Then%>
	<tr>
		<td colspan="3" align="center"><b><%=langRenameTableAlt%></b></td>
	</tr>
<%	end if%>
	<tr>
		<td><%=langNewTableName%></td>
		<td><input type="text" name="tablename" value="<%if action = "rename" Then Response.Write Request.QueryString("table").Item%>"></td>
		<td><input type="submit" name="submit" value="<%if action = "rename" then Response.Write langRenameTableAlt Else Response.Write langCreateNewTable%>" class="button"></td>
	</tr>
	
<%	If StpProfile.ComponentAvailable("ADOX") Then%>	
	<tr>
		<td colspan="3" align="center"><a href="linked.asp"><%=langAddLinkedTable%></a></td>
	</tr>
<%	End If%>	

</table>
</form>
<%
	call DBA_EndNewTable
	set dba = Nothing
%>
<!--#include file=scripts\inc_footer.inc -->
</body>
</html>
