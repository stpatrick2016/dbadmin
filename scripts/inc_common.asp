<%Option Explicit%>
<!--#include file=inc_config.asp -->
<!--#include file=inc_protect.asp -->
<!--#include file=inc_lang.asp -->
<!--#include file=inc_profile_engine.asp -->
<!--#include file=inc_functions.asp -->
<!--#include file=inc_engine.asp -->

<%
'********************************************************
'*	Do not change any values below!						*
'********************************************************

	'DBAdmin Version
	Const DBA_VERSION = "2.0"

	'Name of administrator	
	Const DBA_cfgAdminUsername = "admin"

	'load profile
	call DBA_LoadProfile
%>