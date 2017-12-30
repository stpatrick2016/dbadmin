<%@ Language=VBScript %>
<HTML>
<HEAD>
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href=default.css rel=stylesheet>
</HEAD>
<BODY>
<!--#include file=inc_config.asp -->
<!--#include file=inc_protect.asp -->
<!--#include file=inc_functions.asp -->
<TABLE WIDTH="100%" ALIGN=center>
	<TR>
		<TD width=180 valign=top><!--#include file=inc_nav.asp --></TD>
		<TD>
		
<H1><%=langFreeTypeQuery%></H1>
<P align=center><%=langFreeTypeQueryAlt%> 
</P>
<%
dim rec, con, script, intRecordsAffected, fld, abspage, i, query
dim pagesize
script = Request.ServerVariables("SCRIPT_NAME")
query = Request("query")
pagesize = 10
if Request("pagesize").Count > 0 and IsNumeric(Request("pagesize")) then pagesize = CInt(Request("pagesize"))
if pagesize < 1 then pagesize = 10


if Len(query) > 0 then
	On Error Resume Next
	OpenConnection con
	IsError
	con.CursorLocation = adUseClient
	set rec = con.Execute (query, intRecordsAffected)
		
	if Err = 0 then
%>
<H3 align=center>Total records affected: <B><%=intRecordsAffected%></B></H3>
<%		if rec.State <> adStateClosed then
			rec.CacheSize = pagesize
			rec.PageSize = pagesize
			if rec.PageCount > 0 then
				if Request("page").Count = 0 or CInt(Request("page")) = 0  or rec.PageCount < CInt(Request("page")) then
					rec.AbsolutePage = 1
				else
					rec.AbsolutePage = CInt(Request("page"))
				end if
			end if
			abspage = rec.AbsolutePage
%>
<H3 align=center><%=langTotalRecords%>&nbsp;<B><%=rec.RecordCount%></B></H3>	
<%if rec.RecordCount > 0 then%>
<P align=center>
*&nbsp;<img src="images/xml.gif" border="0" WIDTH="16" HEIGHT="16"><a href="export_xml.asp?sql=<%=Server.URLEncode(query)%>" alt="<%=langXMLExportAlt%>"><%=langXMLExport%></a>&nbsp;
*&nbsp;<img src="images/excel.gif" border="0" WIDTH="16" HEIGHT="16"><a href="export_csv.asp?sql=<%=Server.URLEncode(query)%>" alt="<%=langExcelExporAltt%>"><%=langExcelExport%></a>&nbsp;*
<%end if%>
<table align=center ID="Table1">
	<tr><td align=center>
	<form action="ftquery.asp" method=post>
		<%=langPageSize%>&nbsp;
		<select name="pagesize">
			<option value="5">5</option>
			<option value="10">10</option>
			<option value="15">15</option>
			<option value="25">25</option>
			<option value="50">50</option>
		</select>
		<input type=hidden name="query" value="<%=query%>">
		<input type=submit value="<%=langSubmit%>" class="button">
	</form>
	</td></tr>
</table>
</P>
	<p align="left">
	<%if abspage > 1 then%>
		<a href="<%=script%>?query=<%=Server.URLEncode(query)%>&amp;page=<%=(abspage - 1)%>&amp;pagesize=<%=Request("pagesize")%>"><font size="1">&laquo;&nbsp;<%=langPrev%></font></a>
	<%end if%>
	<%for i=1 to rec.PageCount
		if i = abspage then%>
			<font size="2">[<%=i%>]</font>&nbsp;
	<%	else%>
			<font size="1">&nbsp;[<a href="<%=script%>?query=<%=Server.URLEncode(query)%>&amp;page=<%=i%>&amp;pagesize=<%=Request("pagesize")%>"><%=i%></a>]&nbsp;</font>
	<%	end if
	Next
	if abspage < rec.PageCount and abspage > 0 then%>
		<a href="<%=script%>?query=<%=Server.URLEncode(query)%>&amp;page=<%=(abspage + 1)%>&amp;pagesize=<%=Request("pagesize")%>"><font size="1"><%=langNext%>&nbsp;&raquo;</font></a>
	<%end if
	i = 0
	%>
	</p>

<table align="center" border="1" width="100%">
<tr>
<%for each fld in rec.Fields%>
	<th><%=fld.Name%></th>
<%next%>
</tr>

<%do while not rec.EOF and i < rec.PageSize%>
<tr onmouseover="bgColor='#DDDDDD'" onmouseout="bgColor=''">
	<%for each fld in rec.Fields%>
		<td valign="top" align="center">
		<%if fld.Type <> adBinary then%>
			<%if fld.Value <> "" then%>
				<%=Replace(fld.Value, "<", "&lt;")%>
			<%else%>
				&nbsp;
			<%end if%>
		<%else%>
			&lt;Binary data&gt;
		<%end if%>
		</td>
	<%next%>
</tr>
<%	rec.MoveNext
	i = i + 1 
loop%>

</table>		

	<p align="left">
	<%if abspage > 1 then%>
		<a href="<%=script%>?query=<%=Server.URLEncode(query)%>&amp;page=<%=(abspage - 1)%>&amp;pagesize=<%=Request("pagesize")%>"><font size="1">&laquo;&nbsp;<%=langPrev%></font></a>
	<%end if%>
	<%for i=1 to rec.PageCount
		if i = abspage then%>
			<font size="2">[<%=i%>]</font>&nbsp;
	<%	else%>
			<font size="1">&nbsp;[<a href="<%=script%>?query=<%=Server.URLEncode(query)%>&amp;page=<%=i%>&amp;pagesize=<%=Request("pagesize")%>"><%=i%></a>]&nbsp;</font>
	<%	end if
	Next
	if abspage < rec.PageCount and abspage > 0 then%>
		<a href="<%=script%>?query=<%=Server.URLEncode(query)%>&amp;page=<%=(abspage + 1)%>&amp;pagesize=<%=Request("pagesize")%>"><font size="1"><%=langNext%>&nbsp;&raquo;</font></a>
	<%end if%>
	</p>
<%		end if
	else
		Response.Write "<DIV align=center class=Error>" & Err.Description & "</DIV>"
	end if
end if
%>


      <P align=center><%=langTypeSQL%></P>
      <FORM id=FORM1 name=FORM1 action="<%=script%>" method=post>
      <P ALIGN=CENTER><TEXTAREA id=query name=query rows=5 cols=50><%=query%></TEXTAREA></P>
      <P align=center><INPUT class=button id=submit1 type=submit value="<%=langRunIt%>" name=submit></P>
      </FORM>
		
		</TD>
	</TR>
</TABLE>


</BODY>
</HTML>
