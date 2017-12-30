<table align="left">
<%if Session("DBAdminPassword") = CStr(strAdminPassword) then%>
	<tr>
		<td><img src="images/msaccess.gif" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a HREF="database.asp"><%=langDatabase%></a></td>
	</tr>
	<%if Len(Session("DBAdminDatabase")) > 0 then%>
	<tr>
		<td><img src="images/tables.gif" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a HREF="tablelist.asp"><nobr><%=langTablesList%></nobr></a></td>
	</tr>
	<tr>
		<td><img src="images/query.gif" alt="Stored Procedures List" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a href="qlist.asp"><nobr><%=langProcedures%></nobr></a></td>
	</tr>
	<tr>
		<td><img src="images/view.gif" alt="Views List" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a href="vlist.asp"><nobr><%=langViews%></nobr></a></td>
	</tr>
	<tr>
		<td><img src="images/structure.gif" alt="Relations" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a href="relations.asp"><nobr><%=langRelations%></nobr></a></td>
	</tr>
	<tr>
		<td><img src="images/ftquery.gif" alt="Free-type query" border="0" WIDTH="16" HEIGHT="16">&nbsp;<a href="ftquery.asp"><nobr><%=langFreeTypeQuery%></nobr></a></td>
	</tr>
	<%end if%>
<%else%>
	<tr>
		<td><img src="images/link.gif" alt="<%=langVisitStpWorks%>" border="0" WIDTH="16" HEIGHT="16"><a HREF="http://www.stpworks.com" target="_blank" title="<%=langVisitStpWorks%>"><%=langStPWorks%></a></td>
	</tr>
	<tr>
		<td><img src="images/link.gif" alt="<%=langVisitDBAdmin%>" border="0" WIDTH="16" HEIGHT="16"><a HREF="http://www.stpworks.com/asp/dbadmin.asp" target="_blank" title="<%=langVisitDBAdmin%>"><nobr><%=langDatabaseAdministration%></nobr></a></td>
	</tr>
	<tr>
		<td><img src="images/link.gif" alt="<%=langCheckUpdate%>" border="0" WIDTH="16" HEIGHT="16"><a HREF="http://www.stpworks.com/asp/dbadmin_check.asp?version=<%=Server.URLEncode(cfgStpDBAdminVersion)%>"><%=langCheckUpdate%></a></td>
	</tr>
<%end if%>
</table>
