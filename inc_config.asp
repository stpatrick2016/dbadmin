<!--#include file=inc_lang.asp -->
<!--#include file=config.asp -->

<%
	Response.Buffer = True	'fix for Windows NT hosts
	Response.CacheControl = "no-cache"
	
	'Version number. Please do not change this value!
	Const cfgStpDBAdminVersion = "1.7"
%>