<%@ Language=VBScript %>
<!--#include file=inc_config.asp -->
<html>
<head>
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet">
<title>Database Administration</title>
</head>
<body>
<!--#include file=inc_protect.asp -->
<!--#include file=inc_functions.asp -->
<table WIDTH="100%" ALIGN="center">
	<tr>
		<td width="180" valign="top"><!--#include file=inc_nav.asp --></td>
		<td>
      <h1><%=langTablesList%></h1>
      <p>
<%
	On Error Resume Next
	dim con, rec
	OpenConnection con
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
          <th class=RegularTH><%=langTableName%></th>
          <th colSpan="4" class=RegularTH><%=langAction%></th></tr>
<%	do while not rec.EOF and Err=0%>
        <tr onmouseover="bgColor='#DDDDDD'" onmouseout="bgColor=''">
          <td class=RegularTD><%=rec("Table_name")%></td>
          <td align="center" class=RegularTD><img src="images/structure.gif" alt="<%=langViewTableStructureAlt%>" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a HREF="structure.asp?table=<%=Server.URLEncode(rec("Table_name"))%>"><%=langViewTableStructure%></a></td>
          <td align="center" class=RegularTD><img src="images/table.gif" alt="<%=langViewTableDataAlt%>" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a href="data.asp?table=<%=Server.URLEncode(rec("TABLE_NAME"))%>"><%=langViewTableData%></a></td>
          <td align="center" class=RegularTD><img src="images/cycle.gif" alt="<%=langTableNavigateAlt%>" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a href="recedit.asp?action=edit&amp;table=<%=Server.URLEncode(rec("TABLE_NAME"))%>"><%=langTableNavigate%></a></td>
          <td align="center" class=RegularTD><img src="images/delete.gif" alt="<%=langDeleteTableAlt%>" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a HREF="tablelist.asp?delete=1&amp;table_name=<%=Server.URLEncode(rec("Table_name"))%>" onclick="return confirm('<%=Replace(langSureToDeleteTable, "$table_name", rec("Table_name"))%>?');"><%=langDeleteTable%></a></td></tr>
<%		rec.MoveNext
	Loop%>
       </table></p>

<h3 align="center"><%=langCreateNewTable%></h3>
<p align="center">
<form action="tablelist.asp" method="POST" id="form1" name="form1">
<%=langNewTableName%><input type="text" name="name"><br>
<input type="submit" name="submit" value="<%=langCreateNewTable%>" class="button">
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
