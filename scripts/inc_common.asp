<%Option Explicit%>
<!--#include file=inc_config.asp -->
<!--#include file=inc_protect.asp -->
<!--#include file=inc_lang.asp -->
<!--#include file=inc_profile_engine.asp -->
<!--#include file=inc_LangEngine.asp -->
<!--#include file=inc_functions.asp -->
<!--#include file=inc_engine.asp -->

<%
'********************************************************
'*	Do not change any values below!						*
'********************************************************

	'DBAdmin Version
	Const DBA_VERSION = "2.3"

	'Name of administrator	
	Const DBA_cfgAdminUsername = "admin"
	
	'Session timeout. 0 is default
	Dim DBA_cfgSessionTimeout : DBA_cfgSessionTimeout = 0

	'Configuration object. Loaded in LoadProfile and freed in inc_footer.inc
	Dim StpProfile
	
	'load profile and language
	call DBA_LoadProfile()
	call DBA_LoadLanguage()
	
	'/-----------------------------------------------------------
	'| Addons are under design, they will go here for now
	'| Later I will think how to incorporate addons and plugins
	'\-----------------------------------------------------------
	Dim DBA_cfgAddonsFolder : DBA_cfgAddonsFolder = "plugins"
	Dim DBA_addTextEditor : DBA_addTextEditor = "htmleditor/htmleditor.html"
%>