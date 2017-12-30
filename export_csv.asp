<%@ Language=VBScript %>
<!--#include file=config.asp -->
<!--#include file=inc_protect.asp -->
<!--#include file=inc_functions.asp -->


<%if Request.Form("submit").Count > 0 then%>


<%
'	On Error Resume Next
	dim con, rec, s, fld
	dim DlmColumn, DlmRow, DlmText
	
	if Request.Form("column") = "TAB" then 
		DlmColumn = vbTab
	elseif Request.Form("column") = "SPACE" then 
		DlmColumn = " "
	elseif Request.Form("column") = "OTHER" then 
		DlmColumn = Request.Form("other")
	else
		DlmColumn = Request.Form("column")
	End if
	DlmRow = vbCrLf
	DlmText = Request.Form("text")
	
	set con = Server.CreateObject("ADODB.Connection")
	con.Open strProvider & Session("DBAdminDatabase")
	IsError
	set rec = Server.CreateObject("ADODB.Recordset")
	rec.CursorLocation = adUseClient
	rec.Open Request.Form("sql").Item, con
	
	Response.AddHeader "Content-Disposition", "attachment; filename=" & Session.SessionID & "_export.csv"
	Response.ContentType = "application/octet-stream"
	
	
	'Export field names
	if Request.Form("nofields") <> "1" then
		s = ""
		for each fld in rec.Fields 
			s = s & fld.Name & DlmColumn
		next
		if Len(s) > 0 then s = Left(s, Len(s)-1) & DlmRow
		Response.Write s
	end if

	do while not rec.EOF 
		s = ""
		for each fld in rec.Fields
			Select Case fld.Type 
				Case adBSTR,adChar,adLongVarChar,adLongVarWChar,adVarChar,adVarWChar,adWChar
					s = s & DlmText & Replace(fld.Value, DlmText, DlmText & DlmText) & DlmText & DlmColumn
				Case adBinary, adLongVarBinary, adVarBinary
					s = s & DlmColumn
				Case Else
					s = s & fld.Value & DlmColumn
			End Select
		next
		if Len(s) > 0 then s = Left(s, Len(s)-1) & DlmRow
		Response.Write s
		
		rec.MoveNext
	loop
	Response.Write DlmRow

	rec.Close
	con.Close
	set rec = nothing
	set con = nothing
%>


<%else%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href=default.css rel=stylesheet>
<TITLE>Database Administration</TITLE>
</HEAD>
<BODY>
<TABLE WIDTH=100% ALIGN=center>
	<TR>
		<TD width=180px valign=top><!--#include file=inc_nav.asp --></TD>
		<TD>
		
<H1 align=center>Export as text delimited CSV file</H1>
<P align=center>Please define the column and rows delimiters, or use default and 
click Export</P>
<FORM action="export_csv.asp" method=POST id=form1 name=form1>
<INPUT type=hidden name=sql value="<%=Request.QueryString("sql")%>">
<TABLE ALIGN=center CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD>Column Delimiter</TD>
		<TD>
			<SELECT name="column">
				<OPTION value="TAB">{Tab}</OPTION>
				<OPTION value="SPACE">{Space}</OPTION>
				<OPTION value=";">;</OPTION>
				<OPTION value=",">,</OPTION>
				<OPTION value="OTHER">{Other} --></OPTION>
			</SELECT>&nbsp;
			<INPUT type="text" name="other" size=2 maxlength=1>
		</TD>
	</TR>
	<TR>
		<TD>Text qualifier</TD>
		<TD>
			<SELECT name="text">
				<OPTION value='"'>"</OPTION>
				<OPTION value="'">'</OPTION>
				<OPTION value="">{None}</OPTION>
			</SELECT>
		</TD>
	</TR>
	<TR>
		<TD colspan="2"><INPUT type=checkbox name="nofields" value="1">&nbsp;No field names</TD>
	</TR>
</TABLE>
<P align=center><INPUT type=submit name=submit value="Export" class="button"></P>
</FORM>
		
		</TD>
	</TR>
</TABLE>

<P>&nbsp;</P>

</BODY>
</HTML>
<%end if%>