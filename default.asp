<%@ Language=VBScript %>
<!--#include file=scripts/inc_common.asp-->
<%
	If Request.QueryString("action") = "logoff" Then 
		Session.Contents.Remove(DBA_cfgSessionPwdName)
		Session.Contents.Remove(DBA_cfgSessionUserName)
		Session.Contents.Remove(DBA_cfgSessionDBPathName)
		Session.Contents.Remove(DBA_cfgSessionDBPassword)
	End if
	if Request.Form("password") = DBA_cfgAdminPassword or not DBA_IsSecurityEnabled() then
		Session(DBA_cfgSessionPwdName) = DBA_cfgAdminPassword
		Session(DBA_cfgSessionUserName) = DBA_cfgAdminUsername
		If DBA_cfgSessionTimeout > 0 Then Session.Timeout = DBA_cfgSessionTimeout
		if DBA_IsSecurityEnabled() Then Response.Redirect "database.asp"
	end if
%>
<html>
<head>
<title>DBA:<%=langCaptionHome%></title>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet" type="text/css">
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
<script type="text/javascript" language="javascript">
function onLoad(){
	var obj = document.getElementById('iPassword');
	if(obj){
		obj.focus();
	}
}
function bugReport(){
	var url = 'http://www.stpworks.com/redir.asp?linkid=5&p=1';
	url += '&iis=<%=Server.UrlEncode(GetIISVersion())%>&ado=<%=Server.UrlEncode(GetADOVersion())%>';
	url += '&browser=<%=Server.UrlEncode(Request.ServerVariables("HTTP_USER_AGENT"))%>';
	url += '&se=<%=Server.UrlEncode(ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion)%>';
	DBA_popupWindow(url, 'bug', 640, 480);
}
</script>
</head>

<body onload="javascript:onLoad();">

<%call DBA_WriteNavigation%>
<h2 align="center"><%=langWelcomeHeader%></h2>
<p align="center"><%=langVersion & "&nbsp;" & DBA_VERSION%></p>

<%call DBA_BeginNewTable(langWelcome, "", "75%", "")%>

<p align="center"><%=langWelcomeNote%></p>
<br>

<%
	if Session(DBA_cfgSessionPwdName) <> DBA_cfgAdminPassword and DBA_IsSecurityEnabled() then
		call WriteLoginForm
	else
		call WriteMainPage
	end if
%>

<%call DBA_EndNewTable%>

<!--#include file=scripts/inc_footer.inc-->
</body>
</html>

<%Sub WriteLoginForm%>
<p align="center"><%=langWelcomeNote2%></p>
<form action="default.asp" method="post">
	<table cellspacing="1" cellpadding="1" border="0" align="center">
		<tr>
			<td><%=langEnterPassword%></td>
			<td><input type="password" name="password" id="iPassword"></td>
		</tr>
		<tr>
			<td colspan="2" align="center"><input type="submit" value="<%=langSubmit%>" name="submit" class="button"></td>
		</tr>
	</table>
</form>
<%End Sub%>

<%Sub WriteMainPage%>
<table align="center" border="0" cellpadding="25" cellspacing="1">
	<tr>
		<td align="center" valign="top">
			<a href="settings.asp"><img src="images/icon_settings.gif" border="0" width="48" height="48" alt="<%=langSettings%>"></a>
			<h5><%=langSettings%></h5>
		</td>
		<td align="center" valign="top">
			<a href="javascript:bugReport();">
				<img src="images/icon_submit_bug.gif" width="48" height="48" border="0" alt="<%=langSubmitBug%>">
			</a>
			<h5><%=langSubmitBug%></h5>
		</td>
		<td align="center" valign="top">
			<a href="default.asp?action=logoff"><img src="images/icon_logoff.gif" border="0" width="48" height="48" alt="<%=langLogOff%>"></a>
			<h5><%=langLogOff%></h5>
		</td>
	</tr>
</table>
<%End Sub%>

<%
	'Several functions for bug reporting
	Function GetADOVersion()
		'On Error Resume Next
		dim con
		Set con = Server.CreateObject("ADODB.Connection")
		GetADOVersion = "" & con.Version
		set con = Nothing
	End Function
	
	Function GetIISVersion()
		GetIISVersion = Request.ServerVariables("SERVER_SOFTWARE")
	End Function
%>
