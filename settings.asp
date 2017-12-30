<%@ Language=VBScript %>
<!--#include file=scripts/inc_common.asp-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>DBA:Settings</title>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta name="GENERATOR" content="Microsoft Visual Studio.NET 7.0">
<link href="default.css" rel="stylesheet" type="text/css">
</head>
<body>

<%	call DBA_WriteNavigation%>

<%
	DBA_BeginNewTable langSettings, "", "90%"
	
	if Len(DBA_cfgProfilePath) > 0 then
		if Request.Form("action") = "update" then 
			Session(Request.Form("s_user").Item) = Session(DBA_cfgSessionUserName)
			Session(Request.Form("s_upwd").Item) = Session(DBA_cfgSessionPwdName)
			Session(Request.Form("s_dbpath").Item) = Session(DBA_cfgSessionDBPathName)
			Session(Request.Form("s_dbpwd").Item) = Session(DBA_cfgSessionDBPassword)

			DBA_cfgSessionUserName = Request.Form("s_user").Item
			DBA_cfgSessionPwdName = Request.Form("s_upwd").Item
			DBA_cfgSessionDBPathName = Request.Form("s_dbpath").Item
			DBA_cfgSessionDBPassword = Request.Form("s_dbpwd").Item
			DBA_cfgSaveDBPaths = CBool(Request.Form("save_paths").Item)

			call DBA_SaveProfile
			call DBA_WriteSuccess(langSaveSuccess)
		end if
		call DBA_LoadProfile
%>

<form method="post" action="settings.asp">
<input type="hidden" name="action" value="update">
<table align="center" border="0">
<%	if Session(DBA_cfgSessionUserName) = DBA_cfgAdminUsername then%>
	<tr>
		<th colspan="2"><%=langSessionVariables%></th>
	</tr>
	<tr>
		<td><%=langUsername%></td>
		<td><input type="text" name="s_user" value="<%=DBA_cfgSessionUserName%>"></td>
	</tr>
	<tr>
		<td><%=langUserPassword%></td>
		<td><input type="text" name="s_upwd" value="<%=DBA_cfgSessionPwdName%>"></td>
	</tr>
	<tr>
		<td><%=langDBPath%></td>
		<td><input type="text" name="s_dbpath" value="<%=DBA_cfgSessionDBPathName%>"></td>
	</tr>
	<tr>
		<td><%=langDBPassword%></td>
		<td><input type="text" name="s_dbpwd" value="<%=DBA_cfgSessionDBPassword%>"</td>
	</tr>
<%	end if%>
	<tr>
		<th colspan="2"><%=langOtherSettings%></th>
	</tr>
	<tr>
		<td><%=langSaveDBPaths%></td>
		<td><select name="save_paths">
			<option value="-1"><%=langYes%></option>
			<option value="0" <%if not DBA_cfgSaveDBPaths then Response.Write " selected "%>><%=langNo%></option>
		</select></td>
	</tr>
	<tr><td align="center" colspan="2">
		<input type="submit" name="submit" value="<%=langSubmit%>" class="button">
		<input type="reset" value="<%=langReset%>" class="button">
	</td></tr>
</table>
</form>

<%	
	else
		Response.Write "<p align=center>" & langSettingsNotAvailable & "</p>"
	end if
	call DBA_EndNewTable
%>

<!--#include file=scripts/inc_footer.inc-->
</body>
</html>

