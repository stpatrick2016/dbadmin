<table align="left">
<%if Session("DBAdminPassword") = strAdminPassword then%>
	<tr>
		<td><img src="images/msaccess.gif" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a HREF="database.asp">Database</a></td>
	</tr>
	<%if Len(Session("DBAdminDatabase")) > 0 then%>
	<tr>
		<td><img src="images/tables.gif" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a HREF="tablelist.asp"><nobr>Tables List</nobr></a></td>
	</tr>
	<tr>
		<td><img src="images/query.gif" alt="Stored Procedures List" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a href="qlist.asp"><nobr>Procedures</nobr></a></td>
	</tr>
	<tr>
		<td><img src="images/view.gif" alt="Views List" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a href="vlist.asp"><nobr>Views</nobr></a></td>
	</tr>
	<tr>
		<td><img src="images/structure.gif" alt="Relations" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a href="relations.asp"><nobr>Relations</nobr></a></td>
	</tr>
	<tr>
		<td><img src="images/ftquery.gif" alt="Free-type query" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a href="ftquery.asp"><nobr>Free-type Query</nobr></a></td>
	</tr>
	<%end if%>
<%else%>
	<tr>
		<td><img src="images/link.gif" alt="Visit StP Works web site!" border="0" WIDTH="16" HEIGHT="16"><a HREF="http://www.stpworks.com" target="_blank" title="Visit StP Works web site!">StP Works</a></td>
	</tr>
	<tr>
		<td><img src="images/link.gif" alt="Visit Database Administrator Homepage!" border="0" WIDTH="16" HEIGHT="16"><a HREF="http://www.stpworks.com/asp/dbadmin.asp" target="_blank" title="Visit Database Administrator Homepage!"><nobr>Database Administrator</nobr></a></td>
	</tr>
<%end if%>
</table>
