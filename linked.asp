<%@ Language=VBScript %>
<!--#include file=scripts\inc_common.asp -->
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<head>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<link href="default.css" rel="stylesheet" type="text/css">
<title>DBA:<%=langCaptionLinkedTable%></title>
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
<script type="text/javascript" language="javascript">
var win;
function browseDB(){
	win = window.open("browse.asp", "browse", "innerHeight=400,height=400,innerWidth=300,width=300,status=no,resizable=no,menubar=no,toolbar=no,center=yes,scrollbars=yes", false);
}
</script>
</head>
<body>

<%	call DBA_WriteNavigation%>

<%
	Dim dba, ext_dbpath, ext_table
	
	call DBA_BeginNewTable(langCaptionLinkedTable, "", "90%", "")
	Set dba = new DBAdmin
	ext_dbpath = Request.Form("path").Item
	ext_table = Request.Form("table").Item
	
	If Len(ext_dbpath) > 0 and Len(ext_table) > 0 Then
		'add table link
		call dba.Connect(Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword))
		if not dba.HasError Then call dba.CreateLinkedTable(ext_dbpath, Request.Form("pwd").Item, ext_table, Request.Form("alias").Item)
	End If
	
	'Show database/table selection form
	call DisplayDatabaseSelect()

	call DBA_EndNewTable
	Set dba = Nothing
%>
<!--#include file=scripts\inc_footer.inc -->
</body>
</html>

<%Sub DisplayDatabaseSelect%>
	<p align="center"><%=langLinkedDatabaseSelect%></p>
	<form method="post" action="linked.asp">
	<table align="center" border="0" align="center">
<%	
	If Len(ext_dbpath) > 0 Then
		dim tbl
		call dba.Connect(ext_dbpath, Request.Form("pwd").Item)
		If dba.HasError Then 
			Response.Write "<tr><td align=center colspan=2>"
			call DBA_WriteError(dba.LastError)
			Response.Write "</td></tr>"
		Else
%>
			<tr>
				<td><%=langTableToLink%></td>
				<td><select name="table">
<%			For Each tbl in dba.Tables.Items
				If not tbl.IsLinked and not tbl.IsSystem Then
%>
					<option value="<%=tbl.Name%>"><%=tbl.Name%></option>
<%				End If
			Next
%>
				</select></td>
			</tr>
			<tr>
				<td><%=langAliasName%></td>
				<td><input type="text" name="alias"></td>
			</tr>
<%	
		End If
	End If
%>
		<tr>
			<td><%=langDatabasePath%></td>
			<td>
				<input type="text" name="path" id="iPath" value="<%=ext_dbpath%>">&nbsp;
				<input type="button" value="<%=langBrowseButton%>" class="button" onclick="javascript:browseDB();">
			</td>
		</tr>
		<tr>
			<td><%=langDatabasePassword%></td>
			<td><input type="password" name="pwd" value="<%=Server.HTMLEncode(Request.Form("pwd"))%>"></td>
		</tr>
		<tr><td align="center" colspan="2">
			<input type="submit" value="<%=langSubmit%>" class="button">
		</td></tr>
	</table>
	</form>
<%End Sub%>

