<%@ Language=VBScript %>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet">
<script LANGUAGE="javascript">
<!--
function deleteTable(tName){
	if(confirm("Are you sure you want to delete table " + tName + "?")){
		document.location.replace("tablelist.asp?delete=1&table_name=" +  tName);
	}
}
//-->
</script>
<title>Database Administration</title>
</head>
<body>
<!--#include file=config.asp -->
<!--#include file=inc_protect.asp -->
<!--#include file=inc_functions.asp -->
<table WIDTH="100%" ALIGN="center">
	<tr>
		<td width="180" valign="top"><!--#include file=inc_nav.asp --></td>
		<td>
      <h1>Tables list</h1>
      <p>
<%
	On Error Resume Next
	dim con, rec
	set con = Server.CreateObject("ADODB.Connection")
	con.Open strProvider & Session("DBAdminDatabase") & Session("DBAdminAuth")
	IsError

	if Request.QueryString("delete").Count > 0 and Request.QueryString("table_name").Count > 0 then
		con.Execute "DROP TABLE [" & Request.QueryString("table_name") & "]", adExecuteNoRecords
	end if
	
	if Request.Form("submit").Count > 0 then
		con.Execute "CREATE TABLE [" & Request.Form("name") & "]", adExecuteNoRecords
	end if

	set rec = con.OpenSchema(20, Array(Empty, empty, empty, "TABLE"))
	IsError
%>
      <table cellSpacing="1" cellPadding="1" width="100%" align="center" border="1" class=RegularTable>
        
        <tr>
          <th class=RegularTH>Table name</th>
          <th colSpan="3" class=RegularTH>Actions</th></tr>
<%	do while not rec.EOF and Err=0%>
        <tr onmouseover="bgColor='#DDDDDD'" onmouseout="bgColor=''">
          <td class=RegularTD><%=rec("Table_name")%></td>
          <td align="center" class=RegularTD><img src="images/structure.gif" alt="View table's structure" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a HREF="structure.asp?table=<%=Server.URLEncode(rec("Table_name"))%>">Structure</a></td>
          <td align="center" class=RegularTD><img src="images/table.gif" alt="View table's data" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a href="data.asp?table=<%=Server.URLEncode(rec("TABLE_NAME"))%>">Data</a></td>
          <td align="center" class=RegularTD><img src="images/delete.gif" alt="Delete the table" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a HREF="javascript:void(0);" onclick="javascript:deleteTable('<%=rec("Table_name")%>');">Delete</a></td></tr>
<%		rec.MoveNext
	Loop%>
       </table></p>

<h3 align="center">Create new table</h3>
<p align="center">
<form action="tablelist.asp" method="POST" id="form1" name="form1">
New table name:&nbsp;<input type="text" name="name"><br>
<input type="submit" name="submit" value="Create" class="button">
</form>
</p>

    </td>
	</tr>
</table>
<%
	rec.close
	con.Close
	set rec = nothing
	set con = nothing
%>
</body>
</html>
