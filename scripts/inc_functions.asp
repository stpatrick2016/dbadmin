<%'#writes out navigation controls
Sub DBA_WriteNavigation
	if Len(Session(DBA_cfgSessionPwdName)) = 0 then
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
		<td align="left" width="106"><a href="default.asp" target="" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='Home';return true"><img alt="Home" border="0" src="images/btn_home.gif" width="106" height="27"></a></td>
		<td align="left" width="106"><a href="database.asp" target="" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='Databases';return true"><img alt="Databases" border="0" src="images/btn_databases.gif" width="106" height="27"></a></td>
		<td align="left" width="106"><a href="tablelist.asp" target="" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='Tables';return true"><img alt="Tables" border="0" src="images/btn_tables.gif" width="106" height="27"></a></td>
		<td align="left" width="106"><a href="qlist.asp" target="" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='Procedures';return true"><img alt="Procedures" border="0" src="images/btn_procedures.gif" width="106" height="27"></a></td>
		<td align="left" width="106"><a href="vlist.asp" target="" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='Views';return true"><img alt="Views" border="0" src="images/btn_views.gif" width="106" height="27"></a></td>
		<td align="left" width="106"><a href="relations.asp" target="" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='Relations';return true"><img alt="Relations" border="0" src="images/btn_relations.gif" width="106" height="27"></a></td>
		<td align="left" width="106"><a href="ftquery.asp" target="" onmouseout="window.status=window.defaultStatus;return true" onmousemove="window.status='Free-Type Query';return true"><img alt="Free-Type Query" border="0" src="images/btn_ftquery.gif" width="106" height="27"></a></td>
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
Sub DBA_BeginNewTable(Title, Details, Width)
	dim TableID, imgID, strID
	strID = DBA_GenerateID(Title)
	TableID = "td_" & strID
	imgID = "img_" & strID	
%>
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
					<td valign="top" align="right"><a href="javascript:newTableToggle('<%=TableID%>', '<%=imgID%>');" onclick="newTableToggle('<%=TableID%>', '<%=imgID%>'); return false;"><img src="images/icon_hide.gif" border="0" width="16" height="16" alt="Hide" id="<%=imgID%>"></a></td>
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
		<td colspan="3" align="right" bgcolor="#EFF0DD">&uarr;<a href="#top">Top</a>&nbsp;&nbsp;</td>
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
Sub DBA_WriteError(s)
	Response.Write "<p align=""center"" class=""error"">" & s & "</p>"
End Sub

Sub DBA_WriteSuccess(s)
	Response.Write "<p align=""center""><font color=green>" & s & "</font></p>"
End Sub

Function DBA_FormatDateTime(v)
	v = CDate(v)
	DBA_FormatDateTime = Year(v) & "-" & Month(v) & "-" & Day(v) & " " & Hour(v) & ":" & Minute(v) & ":" & Second(v)
End Function

Function DBA_GenerateID(v)
	if Len(v) > 0 then
		dim re
		set re = new RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "\W"
		DBA_GenerateID = re.Replace(v, "")
		set re = Nothing
	else
		Randomize
		DBA_GenerateID = "AutoID_" & Int(Rnd() * 10000)
	end if
End Function


'loads and saves profile settings
Sub DBA_LoadProfile
	dim cfg
	set cfg = new StpPrivateProfile
	
	'load common settings first
	cfg.Load DBA_cfgProfilePath, ""
	DBA_cfgSessionUserName = cfg.GetProfileString("settings/session", "s_uname", "DBA_AdminUsername")
	DBA_cfgSessionPwdName = cfg.GetProfileString("settings/session", "s_upwd", "DBA_AdminPassword")
	DBA_cfgSessionDBPathName = cfg.GetProfileString("settings/session", "s_dbpath", "DBA_DatabasePath")
	DBA_cfgSessionDBPassword = cfg.GetProfileString("settings/session", "s_dbpwd", "DBA_DatabasePassword")
	
	'load user-dependent settings
	cfg.Load DBA_cfgProfilePath, Session(DBA_cfgSessionUserName)
	DBA_cfgSaveDBPaths = CBool(cfg.GetProfileNumber("settings", "save_paths", -1))
	
	set cfg = Nothing
End Sub

Sub DBA_SaveProfile
	dim cfg
	set cfg = new StpPrivateProfile
	
	'save common settings first
	cfg.Load DBA_cfgProfilePath, ""
	call cfg.SetValue("settings/session", "s_uname", DBA_cfgSessionUserName)
	call cfg.SetValue("settings/session", "s_upwd", DBA_cfgSessionPwdName)
	call cfg.SetValue("settings/session", "s_dbpath", DBA_cfgSessionDBPathName)
	call cfg.SetValue("settings/session", "s_dbpwd", DBA_cfgSessionDBPassword)
	call cfg.Save
	
	'save user-depending settings
	cfg.Load DBA_cfgProfilePath, Session(DBA_cfgSessionUserName)
	call cfg.SetValue("settings", "save_paths", CLng(DBA_cfgSaveDBPaths))
	call cfg.Save
	
	set cfg = Nothing
End Sub

Sub DBA_AppendDatabase(newPath)
	if not DBA_cfgSaveDBPaths then Exit Sub
	
	dim cfg, arrDatabases, i
	set cfg = new StpPrivateProfile
	
	cfg.Load DBA_cfgProfilePath, Session(DBA_cfgSessionUserName)
	arrDatabases = cfg.GetProfileArray("databases", "")
	
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
		call cfg.SetValue("databases", "", arrDatabases)
		cfg.Save
	end if
	
	set cfg = nothing
End Sub

Function DBA_GetDatabases
	DBA_GetDatabases = Array()
	dim cfg
	set cfg = new StpPrivateProfile
	cfg.Load DBA_cfgProfilePath, Session(DBA_cfgSessionUserName)
	DBA_GetDatabases = cfg.GetProfileArray("databases", "")
	
	set cfg = Nothing
End Function

Sub DBA_RemoveDatabase(path)
	dim cfg, arrDatabases, i, arrNew
	arrNew = Array()
	set cfg = new StpPrivateProfile
	
	cfg.Load DBA_cfgProfilePath, Session(DBA_cfgSessionUserName)
	arrDatabases = cfg.GetProfileArray("databases", "")
	
	'check if the database already exist
	for i=0 to ubound(arrDatabases)
		if arrDatabases(i) <> path then
			Redim Preserve arrNew(ubound(arrNew) + 1)
			arrNew(ubound(arrNew)) = arrDatabases(i)
		end if
	next
	
	call cfg.SetValue("databases", "", arrNew)
	cfg.Save
	
	set cfg = Nothing
End Sub
%>