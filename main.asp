<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href=default.css rel=stylesheet>
<TITLE>Database Administration</TITLE>
<SCRIPT LANGUAGE=javascript>
<!--
function onLoad(){
	if(document.getElementById("password") != null)
		document.getElementById("password").focus();
}
//-->
</SCRIPT>

</HEAD>
<BODY onload="javascript:onLoad();"><!--#include file=config.asp -->
<%
	if Request.Form("password") = strAdminPassword then
		Session("DBAdminPassword") = Request.Form("password")
	end if
	
%>
<TABLE WIDTH="100%" ALIGN=center>
	<TR>
		<TD width=180 valign=top><!--#include file=inc_nav.asp --></TD>
		<TD>
      <H1>Login</H1>
<%if Session("DBAdminPassword") <> strAdminPassword then%> 
	<P align=center>
	Welcome to StP Database Administrator for web-based remote administration of your MS Access
	databases.<br>
	Please enter your administrator password and click Login<br>
	If this is a first time you installed StP DBAdmin, make sure to change the default 
	password to your own.
	</P>
      <FORM action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method=post>
      <P align=center>&nbsp;</P>
      <P align=center>Enter administrator password: <INPUT type=password 
      name=password id=password></P>
      <P align=center><INPUT type=submit value=Login name=submit class=button></P>
      <P align=center></FORM>&nbsp;</P>
<%else%>

<%
	if Request.Form("change_pass").Count > 0 and Len(Request.Form("newpass1")) > 0 then
		if Request.Form("newpass1") <> Request.Form("newpass2") then
			Response.Write "<P class=Error align=center>Passwords mismatch</P>"
		else
			dim fso, f
			set fso = Server.CreateObject("Scripting.FileSystemObject")
			
			'check if file exists and remove read-only
			if fso.FileExists(Server.MapPath("config.asp")) then
				set f = fso.GetFile(Server.MapPath("config.asp"))
				if f.Attributes and 1 then f.Attributes = f.Attributes - 1
				set f = nothing
			end if

			set f = fso.CreateTextFile(Server.MapPath("config.asp"), true)
			str =	"<" & "%" & vbCrLf &_
					"Const strAdminPassword = """ & Replace(Request.Form("newpass1"), """", """""") & """" & vbCrLf &_
					"Const strProvider = """ & strProvider & """" & vbCrLf &_
					"Const strDatabases = """ & strDatabases & """" & vbCrLf &_
					"%" & ">"
			f.WriteLine(str)
			f.close
			set f = nothing
			set fso = nothing
		end if
	end if
%>
<P align=center>You are logged into Database administration<BR>
If case you want to change the password, please type it in below
</P>

<FORM action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method=POST>
<TABLE align=center border=0>
<TR>
	<TD>New password</TD>
	<TD><INPUT type=password name=newpass1></TD>
</TR>
<TR>
	<TD>Re-type new password</TD>
	<TD><INPUT type=password name=newpass2></TD>
</TR>
</TABLE>
<P align=center><INPUT type=submit name=change_pass value="Change password" class=button></P>
</FORM>


<%end if%>
</TD>
	</TR>
</TABLE>



</BODY>
</HTML>
