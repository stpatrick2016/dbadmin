<%@ Language=VBScript %>
<!--#include file=inc_config.asp -->
<html>
<head>
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet" type="text/css">
<title><%=langDatabaseAdministration%></title>
<script LANGUAGE="javascript" type="text/javascript">
<!--
var win;
function browseDB(){
	win = window.open("browse.asp", "browse", "innerHeight=400,height=400,innerWidth=300,width=300,status=no,resizable=no,menubar=no,toolbar=no,center=yes,scrollbars=yes", false);
}
//-->
</script>
</head>
<body>
<!--#include file=inc_protect.asp -->
<!--#include file=inc_functions.asp -->
<%
	On Error Resume Next
	dim script, s, arDatabase, jro, sSQL, fso, sFileName, catalog
	arDatabase = Split(strDatabases, "|")
	script = Request.ServerVariables("SCRIPT_NAME")
	if Request.Form("submit").Count > 0 or Request.QueryString("action") = "delete" then
		dim f, str, bFound
		bFound = false
		set fso = Server.CreateObject("Scripting.FileSystemObject")

		'check if the database exists
		if Request.Form("db") = "0" then 
			sFileName = Request.Form("newdb")
		else
			sFileName = Request.Form("db")
		end if
			
		if not fso.FileExists(sFileName) and Request.QueryString("action") <> "delete" then
			if Request.Form("create") = "1" then
				set catalog = Server.CreateObject("ADOX.Catalog")
				if Right(sFileName, 4) <> ".mdb" then sFileName = sFileName & ".mdb"
				catalog.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sFileName
				if Err then
					Response.Write "<p align=center class=""Error"">" & Err.Description & "</p>"
				else
					Session("DBAdminDatabase") = sFileName
					bFound = True
				end if
				set catalog = nothing
			end if
		else
			bFound = true
			if Request.QueryString("action") <> "delete" then
				Session("DBAdminDatabase") = sFileName
				if Len(Request.Form("username").Item) > 0 then
					Session("DBAdminUserID") = CStr(Request.Form("username"))
				else
					Session("DBAdminUserID") = ""
				end if
				Session("DBAdminDBPassword") = CStr(Request.Form("password"))
			end if
		end if

		if bFound then
			bFound = False
				
			'check if config file exists and remove read-only
			if fso.FileExists(Server.MapPath("config.asp")) then
				set f = fso.GetFile(Server.MapPath("config.asp"))
				if f.Attributes and 1 then f.Attributes = f.Attributes - 1
				set f = nothing
			end if
			
			if Request.Form("db") = "0" or Request.QueryString("action") = "delete" then
				set f = fso.CreateTextFile(Server.MapPath("config.asp"), true)
				if Err then
					Response.Write "<P class=Error align=center>" & langCouldnotSaveConfig & "&nbsp;" & Err.Description & "</P>"
				else
					if Request.QueryString("action") = "delete" then
						str = Replace(strDatabases, Request.QueryString("path"), "")
						str = Replace(str, "||", "|")
						if InStrRev(str, "|") = Len(str) then str = Left(str, Len(str) - 1)
						arDatabase = Split(str, "|")
					else
						for each s in arDatabase
							if Len(s) > 0 then str = str & s & "|"
							if StrComp(s, Request.Form("newdb"), 1) = 0 and Len(Request.Form("newdb")) > 0 then bFound = True
						next
						if not bFound then 
							str = str & Request.Form("newdb")
							Redim Preserve arDatabase(UBound(arDatabase) + 1)
							arDatabase(UBound(arDatabase)) = Request.Form("newdb")
						end if
						if Len(str) > 0 and bFound then str = Left(str, Len(str) - 1)
					end if
					str =	"<" & "%" & vbCrLf &_
							"Const strAdminPassword = """ & strAdminPassword & """" & vbCrLf &_
							"Const strProvider = """ & strProvider & """" & vbCrLf &_
							"Const strDatabases = """ & str & """" & vbCrLf &_
							"%" & ">"
					f.WriteLine(str)
					f.close
					set f = nothing
				end if
				set fso = nothing
			end if
		else
			Response.Write "<p align=center class=""Error"">" & langDatabaseNotExists & "</p>"
		end if
	end if
%>
<table WIDTH="100%" ALIGN="center">
	<tr>
		<td width="180" valign="top"><!--#include file=inc_nav.asp --></td>
		<td>
      <h1><%=langDatabaseSelection%></h1>
      <p align="center"><%=langEnterPath%></p>
      <p align="center"><%=langCurrentDatabase%> <b><%=Session("DBAdminDatabase")%></b></p>
      <form id="FORM1" name="FORM1" action="<%=script%>" method="post">
      <table align="center">
        
        <tr>
          <td onmouseover="bgColor='#99CCCC'" onmouseout="bgColor=''" onclick="document.getElementById('db0').checked=true; return true" style="pointer:hand; cursor: hand">
			<input type="radio" name="db" id=db0 value="0" checked><%=langOtherDatabase%>&nbsp;<input type="text" name="newdb" id="newdb"><input type="button" value="Browse" onclick="javascript:browseDB();" class="button"></td></tr>
		<tr><td><input type=checkbox name="create" value="1">&nbsp;<%=langCreateNew%><font style="font-size:75%"><%=langCreateNewAlt%></font></td></tr>
		
<%
	i = 1
	for each s in arDatabase
		if Len(s) > 0 then
%>
        <tr>
			<td onmouseover="bgColor='#99CCCC'" onmouseout="bgColor=''" onclick="document.getElementById('db<%=i%>').checked=true; return true" style="pointer:hand; cursor: hand">
				<input type="radio" name="db" id="db<%=i%>" value="<%=Replace(s, """","\""")%>"><%=s%>&nbsp;&nbsp;<a href="<%=script%>?action=delete&amp;path=<%=Server.URLEncode(s)%>"><img src="images/delete.gif" alt="<%=langRemovePath%>" border="0" WIDTH="16" HEIGHT="16"></a>
			</td>
		</tr>
<%
			i = i + 1
		end if
	Next
%>
	<tr><td>&nbsp;</td></tr>
	<tr>
		<td><!--Username:&nbsp;<input type="text" name="username" id="password" size="10">&nbsp;&nbsp;-->
			<%=langDatabasePassword%>&nbsp;<input type="password" name="password" id="password" size="10">
		</td>
	</tr>
	<tr><td>&nbsp;</td></tr>
    <tr>
		<td align="center"><input type="submit" name="submit" value="<%=langSubmit%>" class="button"></td>
	</tr>
</table>
</form>
      
<%if Len(Session("DBAdminDatabase")) > 0 then%>

<%	
	if Request.QueryString("action") = "compact" then
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		sFileName = Session("DBAdminDatabase")
		sFileName = Left(sFileName, InStrRev(sFileName, "\")) & fso.GetTempName
		set jro = Server.CreateObject("JRO.JetEngine")
		s = "5"
		if Request.QueryString("type") = "97" then s = "4"
		jro.CompactDatabase		"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Session("DBAdminDatabase"), _
								"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & sFileName & "; Jet OLEDB:Engine Type=" & s
		if Err then
			Response.Write "<p class=""Error"" align=center>" & err.Description & "</p>"
			fso.DeleteFile sFileName
		else
			fso.DeleteFile Session("DBAdminDatabase")
			fso.MoveFile sFileName, Session("DBAdminDatabase")
			Response.Write "<p align=center>" & langDatabaseCompacted & "</p>"
		end if
		set jro = nothing
		set fso = nothing
	end if
	
	if Request.QueryString("action") = "backup" then
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		sFileName = Session("DBAdminDatabase")
		sFileName = Left(sFileName, InStrRev(sFileName, ".")) & "bak"
		fso.CopyFile Session("DBAdminDatabase"), sFileName, True
		if Err then
			Response.Write "<p class=""Error"" align=center>" & Err.Description & "</p>"
		else
			Response.Write "<p align=center>" & langBackupCreated & "</p>"
		end if
		set fso = nothing
	end if
	
	if Request.QueryString("action") = "restore" then
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		sFileName = Session("DBAdminDatabase")
		sFileName = Left(sFileName, InStrRev(sFileName, ".")) & "bak"
		fso.CopyFile sFileName, Session("DBAdminDatabase"), True
		if Err then
				Response.Write "<p class=""Error"" align=center>" & Err.Description & "</p>"
		else
			Response.Write "<p align=center>" & langBackupRestored & "</p>"
		end if
		set fso = nothing
	end if
%>

<H2 align=center><%=langDatabaseOptions%></H2>
<table align=center border=0 cellspacing="1">
	<tr><th colspan="2" align=center><%=langAffectCurrent%></th></tr>
	<tr>
		<td valign=top bgcolor="#bbbbff"><a href="<%=script%>?action=compact"><%=langCompactRepair%></a></td>
		<td bgcolor="#bbbbff"><%=langCompactRepairAlt%></td>
	</tr>
	<tr>
		<td valign=top bgcolor="#bbbbff"><a href="<%=script%>?action=compact&amp;type=97"><%=langCompactRepair97%></a></td>
		<td bgcolor="#bbbbff"><%=langCompactRepair97Alt%></td>
	</tr>
	<tr>
		<td valign=top bgcolor="#bbbbff"><a href="<%=script%>?action=backup"><%=langMakeBackup%></a></td>
		<td bgcolor="#bbbbff"><%=langMakeBackupAlt%></td>
	</tr>
	<tr>
		<td valign=top bgcolor="#bbbbff"><a href="<%=script%>?action=restore"><%=langRestoreBackup%></td>
		<td bgcolor="#bbbbff"><%=langRestoreBackupAlt%></td>
	</tr>
</table>
<%end if%>      



      </td>
	</tr>
</table>


</body>
</html>

