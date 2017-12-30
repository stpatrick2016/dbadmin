<%@ Language=VBScript %>
<!--#include file=inc_config.asp -->
<HTML>
<HEAD>
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href=default.css rel=stylesheet>
<SCRIPT LANGUAGE=javascript>
<!--
function onLoad(){
	if(document.getElementById("password") != null)
		document.getElementById("password").focus();
}
//-->
</SCRIPT>

</HEAD>
<BODY onload="javascript:onLoad();">
<%
'	Session("DBAdminPassword") = strAdminPassword
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
	<%=langWelcomeNote%>
	</P>
      <FORM action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method=post>
      <table align=center border=0>
		<tr>
			<td><%=langEnterPassword%></td>
			<td><INPUT type=password name=password id=password></td>
		</tr>
		<tr><td align=center colspan=2>
			<INPUT type=submit value="<%=langSubmit%>" name=submit class=button>
		</td></tr>
      </FORM>
<%else%>

<%
	if Request.Form("change_pass").Count > 0 and Len(Request.Form("newpass1")) > 0 then
		if Request.Form("newpass1") <> Request.Form("newpass2") then
			Response.Write "<P class=Error align=center>" & langPasswordsMismatch & "</P>"
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
<P align=center><%=langLoggedIn%></P>

<FORM action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method=POST>
<TABLE align=center border=0>
<TR>
	<TD><%=langNewPassword%></TD>
	<TD><INPUT type=password name=newpass1></TD>
</TR>
<TR>
	<TD><%=langRetypeNewPassword%></TD>
	<TD><INPUT type=password name=newpass2></TD>
</TR>
</TABLE>
<P align=center><INPUT type=submit name=change_pass value="<%=langChangePassword%>" class=button></P>
</FORM>


<%end if%>
</TD>
	</TR>
</TABLE>



</BODY>
</HTML>
