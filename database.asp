<%@ Language=VBScript %>
<!--#include file=scripts/inc_common.asp -->
<html>
<head>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet" type="text/css">
<title>DBA:<%=langCaptionDatabase%></title>
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
	On Error Resume Next
	dim dba, action, arrDatabases, i, path, filesize
	action = CStr(Request("action").Item)
	set dba = new DBAdmin
	if Request.Form("submit").Count > 0 then
		if Request.Form("new") = "1" then dba.CreateDatabase Request.Form("path") else dba.Connect Request.Form("path"), Request.Form("password")
		if not dba.HasError then 
			Session(DBA_cfgSessionDBPathName) = dba.DatabasePath
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
if action = "remove_path" then call DBA_RemoveDatabase(Request.QueryString("path").Item)

if Len(Session(DBA_cfgSessionDBPathName)) > 0 then
	call DBA_BeginNewTable(langDatabaseOptions, "", "75%", "")
	
	Select Case action
		Case "compact"
			call dba.CompactDatabase(Request.QueryString("upgrade") = "1", null, null)
			if not dba.HasError then DBA_WriteSuccess langDatabaseCompacted
		Case "backup"
			call dba.BackupDatabase
			if not dba.HasError then DBA_WriteSuccess langBackupCreated
		Case "restore"
			call dba.RestoreDatabase
			if not dba.HasError then DBA_WriteSuccess langBackupRestored
		Case "update_password"
			if Request.Form("password").Item = Request.Form("password2").Item then
				call dba.CompactDatabase(False, CStr(Request.Form("password").Item), null)
				if not dba.HasError then 
					DBA_WriteSuccess langNewPasswordSet
					Session(DBA_cfgSessionDBPassword) = CStr(Request.Form("password").Item)
				else
					DBA_WriteError dba.LastError
				end if
			else
				DBA_WriteError langPasswordsMismatch
			end if
		Case "set_lcid"
				call dba.CompactDatabase(False, null, Request.Form("lcid").Item)
	End Select
	if dba.HasError Then call DBA_WriteError(dba.LastError)
	
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
		<tr>
			<td><b><%=langLocaleIdentifier%></b></td>
			<td><%=GetLocaleName(dba.LocaleIdentifier)%></td>
		</tr>
		<tr>
			<td><b><%=langDatabaseType%></b></td>
			<td><%if dba.IsAccess97 then Response.Write "Access 97" else Response.Write "Access 2000"%></td>
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
		<tr><td align="center"><a href="import_db.asp" title="<%=langImportDatabaseAlt%>"><%=langImportDatabase%></a></td></tr>
		<tr><td align="center"><a href="database.asp?action=new_lcid" title="<%=langChangeLocaleIDAlt%>"><%=langChangeLocaleID%></a></td></tr>
<%		if action = "new_lcid" Then%>
			<tr><td align="center">
			<form action="database.asp" method="post">
			<input type="hidden" name="action" value="set_lcid">
			<table align="center" border="0">
				<tr><td><%=langNewLocaleID%></td><td><select name="lcid"><%call GetLocaleIDOptions%></select></td></tr>
				<tr><td colspan="2" align="center"><input type="submit" value="<%=langChangeLocaleID%>" class="button"></td></tr>
			</table>
			</form>
			</td></tr>
<%		end if%>
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
<%call DBA_BeginNewTable(langDatabaseSelection, langDatabaseSelectionAlt, "75%", "")%>
<p align="center"><%=langEnterPath%></p>

<%if Request.Form("submit").Count > 0 and dba.HasError then DBA_WriteError dba.LastError%>

<form action="database.asp" method="post">
<table align="center" border="0">
	<tr>
		<td><%=langDatabasePath%></td>
		<td>
			<input type="text" name="path" id="iPath">&nbsp;
			<input type="button" value="<%=langBrowseButton%>" class="button" onclick="javascript:browseDB();">
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
	<tr><td align="center" colspan="2"><%=langSelectExistingDatabase%></td></tr>
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
		<td align="center" colspan="2"><input type="submit" name="submit" value="<%=langOpenDatabase%>" class="button"></td>
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

<%
	Function GetLocaleName(lcid)
		Select Case lcid
			Case 1033	GetLocaleName = "General"
			Case 2052	GetLocaleName = "Chinese Punctuation"
			Case 133124	GetLocaleName = "Chinese Stroke Count"
			Case 1028	GetLocaleName = "Chinese Stroke Count (Taiwan)"
			Case 197636	GetLocaleName = "Chinese Bopomofo (Taiwan)"
			Case 1050	GetLocaleName = "Croatian"
			Case 1029	GetLocaleName = "Czech"
			Case 1061	GetLocaleName = "Estonian"
			Case 1036	GetLocaleName = "French"
			Case 66615	GetLocaleName = "Georgian Modern"
			Case 66567	GetLocaleName = "German Phonebook"
			Case 1038	GetLocaleName = "Hungarian"
			Case 66574	GetLocaleName = "Hungarian Technical"
			Case 1039	GetLocaleName = "Icelandic"
			Case 1041	GetLocaleName = "Japanese"
			Case 66577	GetLocaleName = "Japanese Unicode"
			Case 1042	GetLocaleName = "Korean"
			Case 66578	GetLocaleName = "Korean Unicode"
			Case 1062	GetLocaleName = "Latvian"
			Case 1036	GetLocaleName = "Lithuaninan"
			Case 1071	GetLocaleName = "FYRO Macedonian"
			Case 1044	GetLocaleName = "Norwegian/Danish"
			Case 1045	GetLocaleName = "Polish"
			Case 1048	GetLocaleName = "Romanian"
			Case 1051	GetLocaleName = "Slovak"
			Case 1060	GetLocaleName = "Slovenian"
			Case 1034	GetLocaleName = "Spanish (Traditional)"
			Case 3082	GetLocaleName = "Spanish (Spain)"
			Case 1053	GetLocaleName = "Swedish/Finnish"
			Case 1054	GetLocaleName = "Thai"
			Case 1055	GetLocaleName = "Turkish"
			Case 1058	GetLocaleName = "Ukranian"
			Case 1066	GetLocaleName = "Vietnamese"
			Case Else	GetLocaleName = "Unknown"
		End Select
	End Function
%>
<%	Sub GetLocaleIDOptions%>
			<option value="1033">General</option>
			<option value="2052">Chinese Punctuation</option>
			<option value="133124">Chinese Stroke Count</option>
			<option value="1028">Chinese Stroke Count (Taiwan)</option>
			<option value="197636">Chinese Bopomofo (Taiwan)</option>
			<option value="1050">Croatian</option>
			<option value="1029">Czech</option>
			<option value="1061">Estonian</option>
			<option value="1036">French</option>
			<option value="66615">Georgian Modern</option>
			<option value="66567">German Phonebook</option>
			<option value="1038">Hungarian</option>
			<option value="66574">Hungarian Technical</option>
			<option value="1039">Icelandic</option>
			<option value="1041">Japanese</option>
			<option value="66577">Japanese Unicode</option>
			<option value="1042">Korean</option>
			<option value="66578">Korean Unicode</option>
			<option value="1062">Latvian</option>
			<option value="1036">Lithuaninan</option>
			<option value="1071">FYRO Macedonian</option>
			<option value="1044">Norwegian/Danish</option>
			<option value="1045">Polish</option>
			<option value="1048">Romanian</option>
			<option value="1051">Slovak</option>
			<option value="1060">Slovenian</option>
			<option value="1034">Spanish (Traditional)</option>
			<option value="3082">Spanish (Spain)</option>
			<option value="1053">Swedish/Finnish</option>
			<option value="1054">Thai</option>
			<option value="1055">Turkish</option>
			<option value="1058">Ukranian</option>
			<option value="1066">Vietnamese</option>
<%	End Sub%>