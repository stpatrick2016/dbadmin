<%@ Language=VBScript %>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet">
<title>Database Administration</title>
</head>
<body>
<!--#include file=config.asp -->
<!--#include file=inc_protect.asp -->
<!--#include file=inc_functions.asp -->
<table WIDTH="100%" ALIGN="center">
	<tr>
		<td width="180px" valign="top"><!--#include file=inc_nav.asp --></td>
		<td>

<h1>Views List</h1>		
<%
	On Error Resume Next
	dim con, rec, script
	script = Request.ServerVariables("SCRIPT_NAME")
	set con = Server.CreateObject("ADODB.Connection")
	con.Open strProvider & Session("DBAdminDatabase") & Session("DBAdminAuth")
	IsError

	'create a view
	if Request.Form("submit").Count > 0 then
		sSQL = "CREATE VIEW [" & Request.Form("name") & "] AS " & Request.Form("code")
		con.Execute sSQL, adExecuteNoRecords
		if Err then
			Response.Write "<P class=Error>" & Err.Description & "</P>"
		end if
	end if
	
	'delete view
	if Request.QueryString("action") = "delete" then
		sSQL = "DROP VIEW [" & Request.QueryString("v") & "]"
		con.Execute sSQL, adExecuteNoRecords
		if Err then
			Response.Write "<P class=Error>" & Err.Description & "</P>"
		end if
	end if

	set rec = con.OpenSchema(adSchemaViews)
	if Err = 0 then
%>
	
<table class="RegularTable" width="100%" border="1" cellpadding="1" cellspacing="1">
	<tr>
		<th class="RegularTH">Name</th>
		<th class="RegularTH">Code</th>
		<th class="RegularTH">Actions</th>
	</tr>
	
	<%do while not rec.EOF and Err=0%>
	<tr onmouseover="bgColor='#DDDDDD'" onmouseout="bgColor=''">
		<td class="RegularTD" valign="top"><%=rec("TABLE_NAME")%></td>
		<td class="RegularTD"><%=HighlightSQL(Replace(rec("VIEW_DEFINITION"), vbCrLf, "<BR>"))%></td>
		<td class="RegularTD" align="center">
			<a href="ftquery.asp?query=<%=Server.URLEncode("SELECT * FROM [" & rec("TABLE_NAME")) & "]"%>"><img src="images/run.gif" alt="Execute view" border="0" WIDTH="16" HEIGHT="16"></a>&nbsp;
			<a href="javascript:deleteView('<%=Server.URLEncode(rec("TABLE_NAME"))%>');"><img src="images/delete.gif" alt="Delete view" border="0" WIDTH="16" HEIGHT="16"></a>
		</td>
	</tr>
	<%	rec.MoveNext
	  loop%>
	
</table>

<p>	
<form action="<%=script%>" method="POST" id="form1" name="form1">
<table align="center" border="0">
	<tr>
		<th align="center" colspan="2"><font size="4"><b>Create a new view</b></font></th>
	</tr>
	<tr>
		<td>View name:</td>
		<td><input type="text" name="name"></td>
	</tr>
	<tr><td colspan="2" align="center"><b>SQL Statement</b></td></tr>
	<tr><td colspan="2" align="center"><textarea name="code" cols="50" rows="6"></textarea></td></tr>
</table>	
<p align="center"><input type="submit" name="submit" value="Create view" class=button></p>
</form>
</p>


<%	end if%>
		</td>
	</tr>
</table>

<%
	rec.Close
	con.Close
	set rec = nothing
	set con = nothing
%>
</body>
<script LANGUAGE="javascript">
<!--
function deleteView(name){
	if(confirm("Are you sure you want to delete view " + name + "?")){
		document.location.replace("<%=script%>?action=delete&v=" + name);
	}
}
//-->
</script>

</html>
