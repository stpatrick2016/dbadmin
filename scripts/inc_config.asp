<%
	Response.Buffer = True	'fix for Windows NT hosts
'	Response.CacheControl = "no-cache"
	Response.Expires = -1000
	
	'//////////////////////////////////////////////////////
	'// Configuration area
	
	'Your password as administrator. You HAVE to change this value before uploading
	'Database Adminstrator!
	Const DBA_cfgAdminPassword = "admin"
	
	'Either relative or full path to a file where all settings will be stored
	'The folder that contains this file should have write permissions.
	'Leave it blank if you don't want to use Settings feature
	Const	DBA_cfgProfilePath = "dbadmin.xml"
	
	'MSXML ProgID. Leave it blank to auto-select, or place the real ProgID
	'If you set it to ProgID, no further check for availability will be done
	Const	DBA_cfgMSXMLProgID = ""
	
	'By setting the next constant to True you can disable security in DBAdmin.
	'This can be useful if you place DBAdmin in secure folder, which by itself
	'requires a user logon. 
	'To enable security again, you can either set it to False or just delete
	Const	DBA_cfgNoSecurity = False
	
	
	'#####################################
	'# All dynamic options. Can be set from Settings page as well (if you have specified DBA_cfgProfilePath)
	
	'Name of Session variable for username
	Dim DBA_cfgSessionUserName : DBA_cfgSessionUserName = "DBA_AdminUsername"
	
	'Name of Session variable to hold administrator password
	Dim DBA_cfgSessionPwdName : DBA_cfgSessionPwdName = "DBA_AdminPassword"
	
	'Name of Session variable to hold a path to current database
	Dim DBA_cfgSessionDBPathName : DBA_cfgSessionDBPathName = "DBA_DatabasePath"
	
	'Name of the Session variable that holds Password for database
	Dim DBA_cfgSessionDBPassword : DBA_cfgSessionDBPassword = "DBA_DatabasePassword"
	
	'Do save database paths? Note, database password won't be stored
	Dim DBA_cfgSaveDBPaths : DBA_cfgSaveDBPaths = True
	
%>

<%' Type library for ADO. If you are getting any errors, see Readme.html file for fix %>
<!-- METADATA TYPE="typelib" uuid="00000205-0000-0010-8000-00AA006D2EA4"  -->
