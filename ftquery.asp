<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href=default.css rel=stylesheet>
<TITLE>Database Administration</TITLE>
</HEAD>
<BODY>
<!--#include file=config.asp -->
<!--#include file=inc_protect.asp -->
<!--#include file=inc_functions.asp -->
<TABLE WIDTH="100%" ALIGN=center>
	<TR>
		<TD width=180 valign=top><!--#include file=inc_nav.asp --></TD>
		<TD>
		
<H1>Free-type query</H1>
<P align=center>Free-type query allows you to make 
      your own SQL statement and get results from it (if there are any results returned) 
</P>
<%
dim rec, con, script, intRecordsAffected, fld, abspage, i, query
script = Request.ServerVariables("SCRIPT_NAME")
query = Request("query")

if Len(query) > 0 then
	On Error Resume Next
	set con = Server.CreateObject("ADODB.Connection")
	con.Open strProvider & Session("DBAdminDatabase")
	IsError
	con.CursorLocation = adUseClient
	set rec = con.Execute (query, intRecordsAffected)
		
	if Err = 0 then
%>
<H3 align=center>Total records affected: <B><%=intRecordsAffected%></B></H3>
<%		if rec.State <> adStateClosed then
			if Request("pagesize").Count > 0 and IsNumeric(Request("pagesize")) then
				rec.CacheSize = CInt(Request("pagesize"))
				rec.PageSize = CInt(Request("pagesize"))
			else
				rec.CacheSize = 10
				rec.PageSize = 10
			end if
			if rec.PageCount > 0 then
				if Request("page").Count = 0 or CInt(Request("page")) = 0  or rec.PageCount < CInt(Request("page")) then
					rec.AbsolutePage = 1
				else
					rec.AbsolutePage = CInt(Request("page"))
				end if
			end if
			abspage = rec.AbsolutePage
%>
<H3 align=center>Total records returned: <B><%=rec.RecordCount%></B></H3>	
<%if rec.RecordCount > 0 then%>
<P align=center>
*&nbsp;<img src="images/xml.gif" border="0" WIDTH="16" HEIGHT="16"><a href="export_xml.asp?sql=<%=Server.URLEncode(query)%>" alt="Export table content as XML file">XML Export</a>&nbsp;
*&nbsp;<img src="images/excel.gif" border="0" WIDTH="16" HEIGHT="16"><a href="export_csv.asp?sql=<%=Server.URLEncode(query)%>" alt="Export as delimited text file">Excel Export</a>&nbsp;*
<%end if%>
</P>
	<p align="left">
	<%if abspage > 1 then%>
		<a href="<%=script%>?query=<%=Server.URLEncode(query)%>&amp;page=<%=(abspage - 1)%>&amp;pagesize=<%=Request("pagesize")%>"><font size="1">&laquo;&nbsp;Prev</font></a>
	<%end if%>
	<%for i=1 to rec.PageCount
		if i = abspage then%>
			<font size="2">[<%=i%>]</font>&nbsp;
	<%	else%>
			<font size="1">&nbsp;[<a href="<%=script%>?query=<%=Server.URLEncode(query)%>&amp;page=<%=i%>&amp;pagesize=<%=Request("pagesize")%>"><%=i%></a>]&nbsp;</font>
	<%	end if
	Next
	if abspage < rec.PageCount and abspage > 0 then%>
		<a href="<%=script%>?query=<%=Server.URLEncode(query)%>&amp;page=<%=(abspage + 1)%>&amp;pagesize=<%=Request("pagesize")%>"><font size="1">Next&nbsp;&raquo;</font></a>
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
		<a href="<%=script%>?query=<%=Server.URLEncode(query)%>&amp;page=<%=(abspage - 1)%>&amp;pagesize=<%=Request("pagesize")%>"><font size="1">&laquo;&nbsp;Prev</font></a>
	<%end if%>
	<%for i=1 to rec.PageCount
		if i = abspage then%>
			<font size="2">[<%=i%>]</font>&nbsp;
	<%	else%>
			<font size="1">&nbsp;[<a href="<%=script%>?query=<%=Server.URLEncode(query)%>&amp;page=<%=i%>&amp;pagesize=<%=Request("pagesize")%>"><%=i%></a>]&nbsp;</font>
	<%	end if
	Next
	if abspage < rec.PageCount and abspage > 0 then%>
		<a href="<%=script%>?query=<%=Server.URLEncode(query)%>&amp;page=<%=(abspage + 1)%>&amp;pagesize=<%=Request("pagesize")%>"><font size="1">Next&nbsp;&raquo;</font></a>
	<%end if%>
	</p>
<%		end if
	else
		Response.Write "<DIV align=center class=Error>Error occured: " & Err.Description & "</DIV>"
	end if
end if
%>


      <P align=center>Type your SQL statement in a box below.</P>
      <FORM id=FORM1 name=FORM1 action="<%=script%>" method=post>
      <P ALIGN=CENTER><TEXTAREA id=query name=query rows=5 cols=50><%=query%></TEXTAREA></P>
      <P align=center><INPUT class=button id=submit1 type=submit value="Run It" name=submit></P>
      </FORM>
		
		</TD>
	</TR>
</TABLE>


</BODY>
</HTML>
