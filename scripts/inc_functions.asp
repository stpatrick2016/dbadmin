<%
	Dim F_LastTableID : F_LastTableID = 0
%>

<%'#writes out navigation controls
Sub DBA_WriteNavigation
	if Len(Session(DBA_cfgSessionPwdName)) = 0 and DBA_IsSecurityEnabled() then
%>
<!-- TABLE BEFORE LOGIN -->
<a name="top"></a>
<table border="0" cellpadding="0" cellspacing="0" width="375" align="center" style="margin-bottom: 20px">
	<tr>
		<td align="left" width="125"><a href="http://www.stpworks.com/" target="_blank" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='Visit StpWorks Web site';return true"><img alt="Visit StpWorks Web site" border="0" src="images/btn_stpworks.gif" width="125" height="27"></a></td>
		<td align="left" width="125"><a href="http://www.stpworks.com/redir.asp?linkid=2" target="_blank" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='Visit Stp Database Administrator Homepage!';return true"><img alt="Visit Stp Database Administrator Homepage!" border="0" src="images/btn_dbadmin_hp.gif" width="125" height="27"></a></td>
		<td align="left" width="125"><a href="javascript:DBA_popupWindow('http://www.stpworks.com/redir.asp?linkid=6&version=<%=DBA_VERSION%>', '_blank', 480, 400);" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='Check for latest version of Database Administrator';return true"><img alt="Check for latest version of Database Administrator (opens in new window)" border="0" src="images/btn_check_update.gif" width="125" height="27"></a></td>
	</tr>
</table>
<%	
	else
%>
<!-- TABLE AFTER LOGIN -->
<a name="top"></a>
<table border="0" cellpadding="0" cellspacing="0" width="636" align="center" style="margin-bottom: 20px">
	<tr>
		<td align="left" width="106"><a href="default.asp" target="" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='<%=langCaptionHome%>';return true"><img alt="<%=langCaptionHome%>" border="0" src="images/btn_home.gif" width="106" height="27"></a></td>
		<td align="left" width="106"><a href="database.asp" target="" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='<%=langCaptionDatabase%>';return true"><img alt="<%=langCaptionDatabase%>" border="0" src="images/btn_databases.gif" width="106" height="27"></a></td>
		<td align="left" width="106"><a href="tablelist.asp" target="" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='<%=langCaptionTablesList%>';return true"><img alt="<%=langCaptionTablesList%>" border="0" src="images/btn_tables.gif" width="106" height="27"></a></td>
		<td align="left" width="106"><a href="qlist.asp" target="" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='<%=langStoredProceduresList%>';return true"><img alt="<%=langStoredProceduresList%>" border="0" src="images/btn_procedures.gif" width="106" height="27"></a></td>
		<td align="left" width="106"><a href="vlist.asp" target="" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='<%=langCaptionViews%>';return true"><img alt="<%=langCaptionViews%>" border="0" src="images/btn_views.gif" width="106" height="27"></a></td>
		<td align="left" width="106"><a href="relations.asp" target="" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='<%=langRelations%>';return true"><img alt="<%=langRelations%>" border="0" src="images/btn_relations.gif" width="106" height="27"></a></td>
		<td align="left" width="106"><a href="ftquery.asp" target="" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='<%=langCaptionFreeTypeQuery%>';return true"><img alt="<%=langCaptionFreeTypeQuery%>" border="0" src="images/btn_ftquery.gif" width="106" height="27"></a></td>
    </tr>
</table>
<table align="center" border="0">
	<tr>
		<td align="center"><small><b><%=langCurrentDatabase%>&nbsp;</b><i><%=Session(DBA_cfgSessionDBPathName)%></i></small></td>
	</tr>
</table>
<%
	end if
End Sub
%>

<%'#begins a table with new layout
Sub DBA_BeginNewTable(Title, Details, Width, HelpID)
	dim TableID, imgID, strID
	F_LastTableID = F_LastTableID + 1
	strID = F_LastTableID
	TableID = "td_" & strID
	imgID = "img_" & strID	
%>
<script language="javascript" type="text/javascript">
//update language strings
var langHide = "<%=langHide%>";
var langShow = "<%=langShow%>";
</script>

<table align=center width="<%=Width%>" cellpadding=0 cellspacing=0 class="newtable">
	<tr>
		<td background="images/table_title_left.gif" class="title" width="150" colspan="1"><nobr><%=Title%></nobr></td>
		<td width="32"><img src="images/table_title_center.gif" width="32" height="24" border="0" alt=""></td>
		<td background="images/table_title_right.gif" style="width:100%"><img src="images/spacer.gif" width="1" height="24" border="0" alt=""></td>
		<td><img src="images/spacer.gif" width="7" height="1" border="0" alt=""></td>
	</tr>
	<tr>
		<td colspan="3" class="subtitle">
			<table align="center" width="100%" border="0">
				<tr>
					<td valign="top"><%=Details%></td>
					<td valign="top" align="right"><a href="javascript:newTableToggle('<%=TableID%>', '<%=imgID%>');" onclick="newTableToggle('<%=TableID%>', '<%=imgID%>'); return false;"><img src="images/icon_hide.gif" border="0" width="16" height="16" alt="<%=langHide%>" id="<%=imgID%>"></a></td>
				</tr>
			</table>
			
		</td>
		<td width="7" background="images/table_shadow_right.gif"><img src="images/spacer.gif" width="7" height="1" border="0" alt=""></td>
	</tr>
	<tr>
		<td bgcolor="#EFF0DD" colspan="3" class="contents" id="<%=TableID%>">
		
<%
End Sub
%>

<%'#finishes a table with new layout
Sub DBA_EndNewTable
%>
		</td>
		<td width="7" background="images/table_shadow_right.gif"></td>
	</tr>
	<tr>
		<td colspan="3" align="right" bgcolor="#EFF0DD">&uarr;<a href="#top"><%=langTop%></a>&nbsp;&nbsp;</td>
		<td width="7" background="images/table_shadow_right.gif"></td>
	</tr>
	<tr>
		<td background="images/table_shadow_bottom.gif" colspan="3"><img src="images/spacer.gif" width="100%" height="5" border="0" alt=""></td>
		<td width="7"><img src="images/table_shadow_corner.gif" width="7" height="5" border="0"></td>
	</tr>
</table>
<%
End Sub
%>

<%
'#################################################################################
'# writes out a message about errors (red)
Sub DBA_WriteError(s)
	Response.Write "<p align=""center"" class=""error"">" & s & "</p>"
End Sub

'#################################################################################
'# writes out a success message (green)
Sub DBA_WriteSuccess(s)
	Response.Write "<p align=""center""><font color=""green"">" & s & "</font></p>"
End Sub

'#################################################################################
'# Formats date and time into SQL formal representation
Function DBA_FormatDateTime(v)
	v = CDate(v)
	DBA_FormatDateTime = Year(v) & "-" & Month(v) & "-" & Day(v) & " " & Hour(v) & ":" & Minute(v) & ":" & Second(v)
End Function


'#################################################################################
'# loads profile into global variables
Sub DBA_LoadProfile
	if IsEmpty(StpProfile) then set StpProfile = new StpPrivateProfile
	
	'load common settings first
	call StpProfile.Load(DBA_cfgProfilePath, "")
	DBA_cfgSessionUserName = StpProfile.GetProfileString("settings/session", "s_uname", "DBA_AdminUsername")
	DBA_cfgSessionPwdName = StpProfile.GetProfileString("settings/session", "s_upwd", "DBA_AdminPassword")
	DBA_cfgSessionDBPathName = StpProfile.GetProfileString("settings/session", "s_dbpath", "DBA_DatabasePath")
	DBA_cfgSessionDBPassword = StpProfile.GetProfileString("settings/session", "s_dbpwd", "DBA_DatabasePassword")
	DBA_cfgSessionTimeout = StpProfile.GetProfileNumber("settings/session", "s_timeout", 0)
	
	'load user-dependent settings
	call StpProfile.Load(DBA_cfgProfilePath, Session(DBA_cfgSessionUserName))
	DBA_cfgSaveDBPaths = CBool(StpProfile.GetProfileNumber("settings", "save_paths", -1))
	
End Sub

'#################################################################################
'# saves profile to the disk
Sub DBA_SaveProfile
	
	if DBA_cfgAdminUsername = Session(DBA_cfgSessionUserName) then
		'save common settings first
		call StpProfile.Load(DBA_cfgProfilePath, "")
		call StpProfile.SetValue("settings/session", "s_uname", DBA_cfgSessionUserName)
		call StpProfile.SetValue("settings/session", "s_upwd", DBA_cfgSessionPwdName)
		call StpProfile.SetValue("settings/session", "s_dbpath", DBA_cfgSessionDBPathName)
		call StpProfile.SetValue("settings/session", "s_dbpwd", DBA_cfgSessionDBPassword)
		call StpProfile.SetValue("settings/session", "s_timeout", DBA_cfgSessionTimeout)
		call StpProfile.Save
	end if
	
	'save user-depending settings
	call StpProfile.Load(DBA_cfgProfilePath, Session(DBA_cfgSessionUserName))
	call StpProfile.SetValue("settings", "save_paths", CLng(DBA_cfgSaveDBPaths))
	call StpProfile.Save
	
End Sub

'#################################################################################
'# Appends a database path to saved databases list and saves the file
Sub DBA_AppendDatabase(newPath)
	if not DBA_cfgSaveDBPaths then Exit Sub
	
	dim arrDatabases, i
	
	arrDatabases = StpProfile.GetProfileArray("databases", "")
	
	'check if the database already exist
	for i=0 to ubound(arrDatabases)
		if arrDatabases(i) = newPath then
			i = -2
			Exit For
		end if
	next
	
	if i <> -2 then
		Redim Preserve arrDatabases(ubound(arrDatabases) + 1)
		arrDatabases(ubound(arrDatabases)) = newPath
		call StpProfile.SetValue("databases", "", arrDatabases)
		call StpProfile.Save
	end if
	
End Sub

'#################################################################################
'# returns an array of saved databases
Function DBA_GetDatabases
	DBA_GetDatabases = Array()
	DBA_GetDatabases = StpProfile.GetProfileArray("databases", "")
End Function

'#################################################################################
'# removes database from the list
Sub DBA_RemoveDatabase(path)
	call StpProfile.RemoveNode("databases/item[.='" & Replace(path, "\", "\\") & "']")
	StpProfile.Save
End Sub

'#################################################################################
'# returns a string with options for combo box for given range and step
Function DBA_GetComboOptions(intStart, intEnd, intStep, intSelected)
	dim ret, i

	ret = ""
	For i=intStart To intEnd Step intStep
		ret = ret & "<option value=""" & i & """"
		if i = intSelected then ret = ret & " selected"
		ret = ret & ">" & i & "</option>" & vbCrLf
	Next
	DBA_GetComboOptions = ret
End Function

'#################################################################################
'# Loads language
Sub DBA_LoadLanguage
	dim le, langFile
	langFile = StpProfile.GetProfileString("settings", "language-file", "")
	if Len(langFile) > 0 Then
		set le = new CMultiLangEngine
		call le.Load("languages/" & langFile, "lang", False)
		set le = Nothing
	End If
End Sub

'#################################################################################
'# Returns options of languages found in Languages folder
Function GetAvailableLanguages
	dim ret, fso, folder, f, le, selLang
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	set folder = fso.GetFolder(Server.MapPath("languages/"))
	selLang = StpProfile.GetProfileString("settings", "language-file", "")
	
	ret = ""
	if TypeName(folder) <> "Nothing" Then
		set le = new CMultiLangEngine
		For each f in folder.Files
			If lcase(Right(f.Name, 5)) = ".lang" Then
				call le.Load("languages/" & f.Name, "", True)
				if Len(le("language")) > 0 Then
					ret = ret & "<option value=""" & f.Name & """"
					if f.Name = selLang Then ret = ret & " selected "
					ret = ret & ">" & le("language") & " (" & f.Name & ")</option>" & vbCrLf
				End If
			End If
		Next
		set le = Nothing
		set folder = Nothing
	End If

	set fso = Nothing
	GetAvailableLanguages = ret
End Function

'#################################################################################
'# returns an integer value if it can be converted to integer or 0
Function VBParseInt(value)
	if Len(value) > 0 and IsNumeric(value) then
		VBParseInt = CLng(value)
	else
		VBParseInt = 0
	end if
End Function

Function getPagingControl(intCount, intPage, intPageSize, strAdditional)
	dim intPages, strRetVal, i, script, filter
	script = Request.ServerVariables("SCRIPT_NAME")
	filter = Request("filter").Item
	if intPageSize = 0 then intPageSize = 10
	intPages = intCount \ intPageSize
	if (intCount Mod intPageSize) > 0 then intPages = intPages + 1
	strRetVal = "<sc" & "ript type=""text/javascript"" language=""javascript"">" & vbCrLf &_
				"function pagingChangePage(page){" & vbCrLf &_
				"location.href='" & script & "?page='+page+'&pagesize=" & pagesize &_
				"&filter=" & Server.URLEncode(filter) & strAdditional & "';" & vbCrLf &_
				"}</sc" & "ript>" & vbCrLf &_
				"<select name=""page"" onchange=""javascript:pagingChangePage(this.options[this.selectedIndex].value)"">"
	
	for i=1 to intPages
		strRetVal = strRetVal & "<option value=""" & i & """"
		if intPage = i then strRetVal = strRetVal & " selected"
		strRetVal = strRetVal & ">" & i & "</option>" & vbCrLf



	next
	strRetVal = strRetVal & "</select>"
	if Len(strRetVal) > 0 then 
		if intPage > 1 then strRetVal = "<a href=""" & script & "?page=" & intPage -1 &_
										"&amp;pagesize=" & intPageSize &_
										"&amp;filter=" & Server.URLEncode(filter) & strAdditional &_
										""">&laquo;&nbsp;<small>" & langPrev & "</small></a>&nbsp;" & strRetVal
		if intPage < intPages then strRetVal =	strRetVal &_
												"&nbsp;<a href=""" & script & "?page=" & intPage + 1 &_
												"&amp;pagesize=" & intPageSize &_
												"&amp;filter=" & Server.URLEncode(filter) & strAdditional &_
												"""><small>" & langNext & "</small>&nbsp;&raquo;</a>"
		strRetVal = "<p align=""left"">" & strRetVal & "</p>"
	end if
	getPagingControl = strRetVal
End Function

%>