<%@ Language=VBScript %>
<!--#include file=scripts\inc_common.asp -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>DBA:<%=langDatabaseExport%></title>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta name="GENERATOR" content="Microsoft Visual Studio.NET 7.0">
<link href="default.css" rel="stylesheet" type="text/css">
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
</head>
<body>
<%	call DBA_WriteNavigation%>

<%
	dim dba, action
	action = CStr(Request("action").Item)
	set dba = new DBAdmin
	dba.Connect Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword)
	if dba.HasError then DBA_WriteError dba.LastError
	
	Select Case action
		Case "generatesql" call GenerateSQLScript
		Case Else call WriteTablesSelect
	End Select

	set dba = nothing
%>

<!--#include file=scripts\inc_footer.inc -->
</body>
</html>

<%
	Sub WriteTablesSelect
		dim item
		DBA_BeginNewTable langDatabaseExport, langDatabaseExportNote, "90%", ""
%>
		<form action="export_db.asp" method="post">
		<table align="left" border="0" cellspacing="1" cellpadding="10">
			<tr class="evenrow">
				<td><b><%=langTablesList%></b><br><select multiple name="table" size="10">
<%		for each item in dba.Tables.Items%>
					<option value="<%=item.Name%>"><%=item.Name%></option>
<%		next%>
				</select></td>
				
				<td><b><%=langViews%><br></b><select name="view" multiple size="10">
<%		for each item in dba.Views.Items%>
					<option value="<%=item.Name%>"><%=item.Name%></option>
<%		next%>
				</select></td>
				
				<td><b><%=langProcedures%><br></b><select name="procedure" multiple size="10">
<%		for each item in dba.Procedures.Items%>
					<option value="<%=item.Name%>"><%=item.Name%></option>
<%		next%>
				</select></td>
<!--				
				<td valign="top">
					<b><%=langOptions%></b><br>
					<input type="checkbox" name="relations" value="-1">&nbsp;<%=langIncludeRelations%>
				</td>
-->			</tr>
			<tr>
				<td colspan="4">
					<input type="hidden" name="action" value="generatesql">
					<input type="submit" name="submit" value="<%=langGenerateSQLScript%>" class="button">
				</td>
			</tr>
		</table>
		</form>
<%
		call DBA_EndNewTable
	End Sub
%>

<%
	Sub GenerateSQLScript
		dim item, sSQLScript, arr, sqlRelations, rel
		sSQLScript = ""
		sqlRelations = ""
		
		'tables
		arr = Split(Request.Form("table").Item, ",")
		for each item in arr
			item = Trim(item)
			if dba.Tables.Exists(item) then sSQLScript = sSQLScript & dba.Tables.Item(item).SQL & vbCrLf
		next
		
		'views
		arr = Split(Request.Form("view").Item, ",")
		for each item in arr
			item = Trim(item)
			if dba.Views.Exists(item) then sSQLScript = sSQLScript & dba.Views.Item(item).SQL & vbCrLf
		next
		
		'stored procedures
		arr = Split(Request.Form("procedure").Item, ",")
		for each item in arr
			item = Trim(item)
			if dba.Procedures.Exists(item) then sSQLScript = sSQLScript & dba.Procedures.Item(item).SQL & vbCrLf
		next
		
		
		DBA_BeginNewTable langDatabaseExport, langSQLScriptNote, "90%", ""
%>
		<table align="left" border="0">
			<tr>
				<td><textarea cols="60" rows="20" id="taSQL" style="width: 100%"><%=sSQLScript%></textarea></td>
			</tr>
			<tr>
				<td>
					<input type="button" value="<%=langCopyToClipboard%>" onclick="javascript:copyToClipboard(document.getElementById('taSQL'));" class="button">
					<input type="button" value="<%=langCancel%>" onclick="javascript:window.location.href='export_db.asp';" class="button">
				</td>
			</tr>
		</table>
<%
		call DBA_EndNewTable
	End Sub
%>