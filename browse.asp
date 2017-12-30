<%@ Language=VBScript %>
<!--#include file=scripts/inc_common.asp -->
<html>
<head>
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" href="default.css" type="text/css">
<title><%=langBrowse%></title>
<script LANGUAGE="javascript">
<!--
function pickFile(s){
	if(window.opener != null && window.opener != 'undefined'){
		window.opener.document.getElementById("iPath").value = s;
		window.close();
	}else{
		prompt("<%=langOldBrowser%>",s);
		window.close();
	}
}
function onDriveChange(drive){
	if(drive.length > 0){
		var s = drive;
		if(s.lastIndexOf("\\") != s.length-1)
			s += "\\";
		document.location.replace("browse.asp?dir=" + escape(s));
	}
}
//-->
</script>
</head>
<body>
<h3><%=langBrowse%></h3>
<P align=center><%=langOnlyMDB%></P>
<%
dim fso, dir, strCurDir, s, script, d
script = Request.ServerVariables("SCRIPT_NAME")
if Request.QueryString("dir").Count = 0 then 
	strCurDir = Server.MapPath("/")
else
	strCurDir = Request.QueryString("dir")
end if
set fso = Server.CreateObject("Scripting.FileSystemObject")
On Error Resume Next

set dir = fso.GetFolder(strCurDir)
if Err then
	Response.Write "<P class=Error align=center>" & langCannotAccessFolder & Err.Description & "</P>"
	strCurDir = Request.QueryString("curdir")
	set dir = fso.GetFolder(strCurDir)
end if
%>
<P align=left>
	<b><%=langDriveSelector%></B><BR>
	<SELECT name=drive id=drive onchange="javascript:onDriveChange(this.options[this.selectedIndex].value);">
<%
set d = fso.GetDrive(fso.GetDriveName(fso.GetAbsolutePathName(strCurDir)))
%>
		<OPTION value="<%=d.Path%>"><%=d.Path%>&nbsp;(<%=GetDriveType(d.DriveType)%>)</OPTION>
		<OPTION value="">---------</OPTION>
<%for each d in fso.Drives%>
		<OPTION value="<%=d.Path%>"><%=d.Path%>&nbsp;(<%=GetDriveType(d.DriveType)%>)</OPTION>
<%next%>
	</SELECT>
</P>
<table align="left" border="0" width="100%">
<TR><td style="border:1px solid navy"><b><%=langCurrentPath%></B><br><%=strCurDir%></td></tr>
<%
set d = nothing
if not dir.IsRootFolder then
%>
<tr><td><img src="images/folder.gif" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a href="<%=script%>?dir=<%=Server.URLEncode(dir.ParentFolder.Path)%>&amp;curdir=<%=Server.URLEncode(strCurDir)%>">[...]</a></td></tr>
<%end if%>

<%for each s in dir.SubFolders%>
<tr><td><img src="images/folder.gif" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a href="<%=script%>?dir=<%=Server.URLEncode(s.Path & "\")%>&amp;curdir=<%=Server.URLEncode(strCurDir)%>"><%=s.Name%></a></td></tr>
<%next%>

<%for each s in dir.Files%>
	<%if StrComp(Right(s.Name, 3), "mdb", 1) = 0 then%>
		<tr><td><img src="images/msaccess.gif" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a href="javascript:pickFile('<%=Replace(s.Path, "\", "\\")%>');"><%=s.Name%></a></td></tr>
	<%end if%>
<%next%>
</table>

<%
set dir = nothing
set fso = nothing
%>
</body>
</html>
<SCRIPT LANGUAGE=vbscript RUNAT=Server>
Function GetDriveType(dtype)
	Select Case dtype
		Case 1: GetDriveType = langRemovable
		Case 2: GetDriveType = langFixed
		Case 3: GetDriveType = langNetwork
		Case 4: GetDriveType = "CD-ROM"
		Case 5: GetDriveType = "RAM-Disk"
		Case Else:	GetDriveType = langUnknown
	End Select
End Function
</SCRIPT>
