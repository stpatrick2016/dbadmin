<%
if	Session(DBA_cfgSessionPwdName) <> DBA_cfgAdminPassword and _
	InStr(1, Request.ServerVariables("SCRIPT_NAME"), "default.asp", vbTextCompare) <= 0 and _
	DBA_IsSecurityEnabled() _
	then Response.Redirect "default.asp"
	
Function DBA_IsSecurityEnabled
	On Error Resume Next
	If not DBA_cfgNoSecurity = True Then DBA_IsSecurityEnabled = True Else DBA_IsSecurityEnabled = False
	On Error Goto 0
End Function
%>