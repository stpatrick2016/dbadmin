<%@ Language=VBScript %>
<!--#include file=inc_config.asp -->
<html>
<head>
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet">
<SCRIPT LANGUAGE=javascript>
<!--
function runSP(sp){
	var params = window.prompt("<%=langEnterQParams%>", "");
	if(params != null && params.length > 0)
		window.location.reload("ftquery.asp?query=" + escape("EXECUTE [" + sp + "] " + params));
}
//-->
</SCRIPT>
</head>
<body>
<!--#include file=inc_protect.asp -->
<!--#include file=inc_functions.asp -->
<table WIDTH="100%" ALIGN="center">
	<tr>
		<td width="180px" valign="top"><!--#include file=inc_nav.asp --></td>
		<td>

<h1><%=langStoredProceduresList%></h1>		
<%
	on Error Resume Next
	dim con, rec, script, sSQL
	script = Request.ServerVariables("SCRIPT_NAME")
	OpenConnection con
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
		<th class="RegularTH"><%=langSPName%></th>
		<th class="RegularTH"><%=langSPCode%></th>
		<th class="RegularTH"><%=langSPActions%></th>
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
		<th align="center"><font size="4"><b><%=langCreateProcedure%></b></font></th>
	</tr>
	<tr>
		<th align="center"><%=langCreateProcedureNote%></th>
	</tr>
	<tr><td align=center><b><%=langProcedureName%></b></td></tr>
	<tr><td align=center><input type="text" name="name"></td></tr>
	<tr><td align="center"><b><%=langSQLStatement%></b><br>
		<%=langParams1stWay%><br>
		<pre>
		PARAMETERS <EM>Param1</EM> LONG, <EM>Param2</EM> TEXT(255);
		SELECT * FROM Table1 WHERE Column1=<EM>Param1</EM> OR Column2=<EM>Param2</EM>;</pre>
		<%=langParams2ndWay%>
	</td></tr>
	<tr><td align="center"><textarea name="code" cols="50" rows="6"></textarea></td></tr>
</table>	
<p align="center"><input type="submit" name="submit" value="<%=langCreateProcedure%>" class=button></p>
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
	if(confirm("<%=langDeleteProcedurePrompt%> " + name + "?")){
		document.location.replace("<%=script%>?action=delete&q=" + name);
	}
}
//-->
</script>
</html>
