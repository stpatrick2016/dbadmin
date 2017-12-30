<%
if	Session(DBA_cfgSessionPwdName) <> DBA_cfgAdminPassword and _
	InStr(1, Request.ServerVariables("SCRIPT_NAME"), "default.asp", vbTextCompare) <= 0 _
	then Response.Redirect "default.asp"
%>