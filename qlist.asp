<%@ Language=VBScript %>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet">
<title>Database Administration</title>
<SCRIPT LANGUAGE=javascript>
<!--
function runSP(sp){
	var params = window.prompt("Please enter the parameters of procedure divided by commas. Remember to enclose text parameters in single quotation marks.", "");
	if(params != null && params.length > 0)
		window.location.reload("ftquery.asp?query=" + escape("EXECUTE [" + sp + "] " + params));
}
//-->
</SCRIPT>
</head>
<body>
<!--#include file=config.asp -->
<!--#include file=inc_protect.asp -->
<!--#include file=inc_functions.asp -->
<table WIDTH="100%" ALIGN="center">
	<tr>
		<td width="180px" valign="top"><!--#include file=inc_nav.asp --></td>
		<td>

<h1>Stored Procedures (Queries) List</h1>		
<%
	on Error Resume Next
	dim con, rec, script, sSQL
	script = Request.ServerVariables("SCRIPT_NAME")
	set con = Server.CreateObject("ADODB.Connection")
	con.Open strProvider & Session("DBAdminDatabase") & Session("DBAdminAuth")
	IsError
	
	'create a procedure
	if Request.Form("submit").Count > 0 then
		sSQL = "CREATE PROCEDURE [" & Request.Form("name") & "] AS " & Request.Form("code")
		con.Execute sSQL, adExecuteNoRecords
		if Err then
			Response.Write "<P class=Error>" & Err.Description & "</P>"
		end if
	end if
	
	'delete procedure
	if Request.QueryString("action") = "delete" then
		sSQL = "DROP PROCEDURE [" & Request.QueryString("q") & "]"
		con.Execute sSQL, adExecuteNoRecords
		if Err then
			Response.Write "<P class=Error>" & Err.Description & "</P>"
		end if
	end if
	
	set rec = con.OpenSchema(adSchemaProcedures)
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
		<td class="RegularTD" valign="top"><%=rec("PROCEDURE_NAME")%></td>
		<td class="RegularTD"><%=HighlightSQL(Replace(rec("PROCEDURE_DEFINITION"), vbCrLf, "<BR>"))%></td>
		<td class="RegularTD" align="center">
			<a href="javascript:runSP('<%=rec("PROCEDURE_NAME")%>');"><img src="images/run.gif" alt="Execute Stored Procedure" border="0" WIDTH="16" HEIGHT="16"></a>&nbsp;
			<a href="javascript:deleteProcedure('<%=Server.URLEncode(rec("PROCEDURE_NAME"))%>');"><img src="images/delete.gif" alt="Delete procedure" border="0" WIDTH="16" HEIGHT="16"></a>
		</td>
	</tr>
	<%	rec.MoveNext
	  loop%>
	
</table>

<p>	
<form action="<%=script%>" method="POST">
<table align="center" border="0">
	<tr>
		<th align="center"><font size="4"><b>Create a new procedure</b></font></th>
	</tr>
	<tr>
		<th align="center">Note, if you won't add any parameter in your SQL statement, then a new
		View will be created instead</th>
	</tr>
	<tr><td align=center><b>Procedure name:</b></td></tr>
	<tr><td align=center><input type="text" name="name"></td></tr>
	<tr><td align="center"><b>SQL Statement</b><br>
		Parameters can be defined in 2 ways. First way, add PARAMETERS clause in your SQL statement
		with all parameters and thier types listed. For example:<br>
		<pre>
		PARAMETERS <EM>Param1</EM> LONG, <EM>Param2</EM> TEXT(255);
		SELECT * FROM Table1 WHERE Column1=<EM>Param1</EM> OR Column2=<EM>Param2</EM>;</pre>
		The second way, when parameters determined on-the-fly. If you will add a parameter that
		is not recognized as a column name or SQL reserved word, it will be threated as parameter.
	</td></tr>
	<tr><td align="center"><textarea name="code" cols="50" rows="6"></textarea></td></tr>
</table>	
<p align="center"><input type="submit" name="submit" value="Create procedure" class=button></p>
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
function deleteProcedure(name){
	if(confirm("Are you sure you want to delete a stored procedure " + name + "?")){
		document.location.replace("<%=script%>?action=delete&q=" + name);
	}
}
//-->
</script>
</html>
