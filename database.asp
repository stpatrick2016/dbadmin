<%@ Language=VBScript %>
<!--#include file=scripts/inc_common.asp -->
<html>
<head>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet" type="text/css">
<title>DBA:Database</title>
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
<script LANGUAGE="javascript" type="text/javascript">
<!--
var win;
function browseDB(){
	win = window.open("browse.asp", "browse", "innerHeight=400,height=400,innerWidth=300,width=300,status=no,resizable=no,menubar=no,toolbar=no,center=yes,scrollbars=yes", false);
}
function onDatabaseChange(dbpath){
	var obj = document.getElementById('iPath');
	if(obj){
		obj.value = dbpath;
	}
	obj = document.getElementById('cbNew');
	if(obj){
		obj.checked = false;
	}
}
function removeDBPath(){
	var obj = document.getElementById("selDB");
	if(obj){
		window.location.href = "database.asp?action=remove_path&path=" + escape(obj.options[obj.selectedIndex].value);
	}
}
//-->
</script>
</head>
<body>
<%
	dim dba, action, arrDatabases, i, path, filesize
	action = CStr(Request("action").Item)
	set dba = new DBAdmin
	if Request.Form("submit").Count > 0 then
		if Request.Form("new") = "1" then dba.CreateDatabase Request.Form("path") else dba.Connect Request.Form("path"), Request.Form("password")
		if not dba.HasError then 
			Session(DBA_cfgSessionDBPathName) = CStr(Request.Form("path").Item)
			Session(DBA_cfgSessionDBPassword) = CStr(Request.Form("password").Item)
			
			DBA_AppendDatabase CStr(Request.Form("path").Item)
		end if
	elseif Len(Session(DBA_cfgSessionDBPathName)) > 0 then
		dba.Connect Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword)
	end if
%>

<%call DBA_WriteNavigation%>

<!--DATABASE OPTIONS-->
<%
if Len(Session(DBA_cfgSessionDBPathName)) > 0 then
	call DBA_BeginNewTable(langDatabaseOptions, "", "75%")
	
	Select Case action
		Case "compact"
			dba.CompactDatabase Request.QueryString("upgrade") = "1", null
			if not dba.HasError then DBA_WriteSuccess langDatabaseCompacted
		Case "backup"
			call dba.BackupDatabase
			if not dba.HasError then DBA_WriteSuccess langBackupCreated
		Case "restore"
			call dba.RestoreDatabase
			if not dba.HasError then DBA_WriteSuccess langBackupRestored
		Case "update_password"
			if Request.Form("password").Item = Request.Form("password2").Item then
				dba.CompactDatabase False, CStr(Request.Form("password").Item)
				if not dba.HasError then 
					DBA_WriteSuccess langNewPasswordSet
					Session(DBA_cfgSessionDBPassword) = CStr(Request.Form("password").Item)
				else
					DBA_WriteError dba.LastError
				end if
			else
				DBA_WriteError langPasswordsMismatch
			end if
		Case "remove_path"
			call DBA_RemoveDatabase(Request.QueryString("path").Item)
	End Select
	
	if Request.QueryString("action").Count > 0 and dba.HasError then DBA_WriteError dba.LastError
	filesize = dba.Size
%>
	<fieldset>
	<legend><%=langProperties%></legend>
	<table border="0">
		<tr>
			<td><b><%=langFileSize%></b></td>
			<td><%=FormatNumber(filesize, 0, False, False, True)%> bytes</td>
		</tr>
		<tr>
			<td><b><%=langSizeAfterCompact%></b></td>
			<td><%=FormatNumber(filesize - dba.ReclaimedSpace, 0, False, False, True)%> bytes (- <%=FormatNumber(dba.ReclaimedSpace, 0, True, False, True)%> bytes)</td>
		</tr>
	</table>
	</fieldset>
	
	<fieldset title="<%=langActions%>">
	<legend><%=langActions%></legend>
	<p align="center"><%=langAffectCurrent%></p>
	<table align="center" border="0">
<%	if dba.IsAccess97 then%>
		<tr><td align="center"><a href="database.asp?action=compact&amp;upgrade=1" title="<%=langConvert2000Alt%>"><%=langConvert2000%></a></td></tr>
<%	end if%>
		<tr><td align="center"><a href="database.asp?action=compact" title="<%=langCompactRepairAlt%>"><%=langCompactRepair%></a></td></tr>
		<tr><td align="center"><a href="database.asp?action=backup" title="<%=langMakeBackupAlt%>"><%=langMakeBackup%></a></td></tr>
		<tr><td align="center"><a href="database.asp?action=restore" title="<%=langRestoreBackupAlt%>"><%=langRestoreBackup%></a></td></tr>
		<tr><td align="center"><a href="export_db.asp" title="<%=langDatabaseExportAlt%>"><%=langDatabaseExport%></a></td></tr>
		<tr><td align="center"><a href="database.asp?action=newpassword" title="<%=langNewDatabasePassword%>"><%=langNewDatabasePassword%></a></td></tr>
<%		if action = "newpassword" and Request.Form("password").Count = 0 then%>
			<tr><td align="center"><p align="center"><%=langNewDatabasePasswordAlt%></p>
			<form action="database.asp" method="post">
			<input type="hidden" name="action" value="update_password">
			<table align="center" border="0">
				<tr><td><%=langNewPassword%></td><td><input type="password" name="password"></td></tr>
				<tr><td><%=langRetypeNewPassword%></td><td><input type="password" name="password2"></td></tr>
				<tr><td align="center" colspan="2"><input type="submit" name="submit_password" value="<%=langChangePassword%>" class="button"></td></tr>
			</table>
			</form></td></tr>
<%		end if%>
	</table></fieldset>
<%
	call DBA_EndNewTable
end if
%>


<!--DATABASE SELECTION-->
<%call DBA_BeginNewTable(langDatabaseSelection, langDatabaseSelectionAlt, "75%")%>
<p align="center"><%=langEnterPath%></p>

<%if Request.Form("submit").Count > 0 and dba.HasError then DBA_WriteError dba.LastError%>

<form action="database.asp" method="post">
<table align="center" border="0">
	<tr>
		<td><%=langDatabasePath%></td>
		<td>
			<input type="text" name="path" id="iPath">&nbsp;
			<input type="button" value="Browse" class="button" onclick="javascript:browseDB();">
		</td>
	</tr>
	<tr>
		<td><%=langDatabasePassword%></td>
		<td><input type="password" name="password"></td>
	</tr>
	<tr>
		<td align="center" colspan="2"><input type="checkbox" value="1" name="new" id="cbNew" title="<%=langCreateNewAlt%>">&nbsp;<%=langCreateNew%></td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr><td align="center" colspan="2">Select existing database</td></tr>
	<tr>
		<td align="center" colspan="2">
			<select name="db" id="selDB" onchange="javascript:onDatabaseChange(this.options[this.selectedIndex].value);">
				<option value=""></option>
<%	
	arrDatabases = DBA_GetDatabases()
	for i=0 to ubound(arrDatabases)
		path = arrDatabases(i)
		if Len(path) > 15 then
			path =	Left(arrDatabases(i), InStr(4, arrDatabases(i), "\")) & "...\" &_
					Right(arrDatabases(i), Len(arrDatabases(i)) - InStrRev(arrDatabases(i), "\"))
		end if
%>
				<option value="<%=arrDatabases(i)%>"><%=path%></option>
<%	next%>
			</select>&nbsp;<a href="javascript:removeDBPath();"><img src="images/delete.gif" border="0" width="16" height="16" alt="<%=langRemoveDBPathAlt%>"></a>
		</td>
	</tr>
	<tr>
		<td align="center" colspan="2"><input type="submit" name="submit" value="Open database" class="button"></td>
	</tr>
</table>
</form>
<%
	call DBA_EndNewTable
	set dba = Nothing
%>

<!--#include file=scripts/inc_footer.inc-->
</body>
</html>

