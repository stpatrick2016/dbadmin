<!--#include file=adovbs.inc -->

<%
if Session("DBAdminPassword") <> strAdminPassword then Response.Redirect "main.asp"
%>