<%@ Language=VBScript %>
<!--#include file=scripts/inc_common.asp-->
<html>
<head>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet" type="text/css">
<title>DBA:<%=langCaptionFreeTypeQuery%></title>
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
</head>
<body>

<%	call DBA_WriteNavigation%>

<%
	dim dba, strQuery, rec, AffectedRecords, pagesize, page, abspage, sClass, i, fld, strFilter
	strQuery = CStr(Request("query").Item)
	if IsNumeric(Request("pagesize").Item) then pagesize = CInt(Request("pagesize").Item) else pagesize = 10
	if IsNumeric(Request("page").Item) then page = CInt(Request("page").Item) else page = 1
	if page < 1 then page = 1
	if pagesize < 1 then pagesize = StpProfile.GetProfileNumber("settings", "page_size", 10)
	if pagesize < 1 then pagesize = 10

	set dba = new DBAdmin
	dba.Connect Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword)
	if Len(strQuery) > 0 then 
		set rec = dba.RunScript( _
			strQuery, _
			Request.Form("transaction").Item, _
			Request.Form("ignore_errors").Item, _
			AffectedRecords)
	end if
%>

<!--BEGIN RESULTS FORM-->
<%
	if Len(strQuery) > 0 then
		DBA_BeginNewTable langFTQResults, "", "90%", ""
%>
<%		
		'first let's write out what was affected
		for each i in AffectedRecords
			DBA_WriteSuccess langRecordsAffected & "&nbsp;" & i & "<br>"
		next
		if rec.State <> adStateClosed then
			rec.CacheSize = pagesize
			rec.PageSize = pagesize
			if rec.PageCount > 0 then rec.AbsolutePage = page
			abspage = rec.AbsolutePage
%>
<h3 align="center"><%=langTotalRecords%>&nbsp;<b><%=rec.RecordCount%></b></h3>	

	<!--BEGIN EXPORT OPTIONS-->
<%			if rec.RecordCount > 0 then%>
<p align=center>
*&nbsp;<img src="images/xml.gif" border="0" width="16" height="16"><a href="export_xml.asp?sql=<%=Server.URLEncode(strQuery)%>" alt="<%=langXMLExportAlt%>"><%=langXMLExport%></a>&nbsp;
*&nbsp;<img src="images/excel.gif" border="0" width="16" height="16"><a href="export_csv.asp?sql=<%=Server.URLEncode(strQuery)%>" alt="<%=langExcelExportAlt%>"><%=langExcelExport%></a>&nbsp;*
</p>
<%			end if%>
	<!--END EXPORT OPTIONS-->

<table align="center">
	<tr><td align="center">
	<form action="ftquery.asp" method="post">
			<%=langFilter%>&nbsp;
			<select name="filter_field">
				<option value=""></option>
<%	For Each fld in rec.Fields%>
				<option value="<%=fld.Name%>"><%=fld.Name%></option>
<%	Next%>
			</select>
			<select name="filter_cmp">
				<option value="=">=</option>
				<option value=">">></option>
				<option value="<"><</option>
				<option value=">=">>=</option>
				<option value="<="><=</option>
				<option value="<>"><></option>
				<option value="LIKE">LIKE</option>
			</select>
			<input type="text" name="filter_criteria" size="10">
			
			&nbsp;&nbsp;
		<%=langPageSize%>&nbsp;
		<select name="pagesize">
			<%=DBA_GetComboOptions(5, 50, 5, pagesize)%>
		</select>
		<input type="hidden" name="query" value="<%=Replace(strQuery, """", "&quot;")%>">
		<input type="submit" value="<%=langSubmit%>" class="button">
	</form>
	</td></tr>
</table>

	<p align="left">
<%			if abspage > 1 then%>
				<a href="ftquery.asp?query=<%=Server.URLEncode(strQuery)%>&amp;page=<%=(abspage - 1)%>&amp;pagesize=<%=pagesize%>"><font size="1">&laquo;&nbsp;<%=langPrev%></font></a>
<%			end if%>
<%			for i=1 to rec.PageCount
				if i = abspage then%>
					<font size="2">[<%=i%>]</font>&nbsp;
			<%	else%>
					<font size="1">&nbsp;[<a href="ftquery.asp?query=<%=Server.URLEncode(strQuery)%>&amp;page=<%=i%>&amp;pagesize=<%=pagesize%>"><%=i%></a>]&nbsp;</font>
			<%	end if
			Next
			if abspage < rec.PageCount and abspage > 0 then%>
				<a href="ftquery.asp?query=<%=Server.URLEncode(strQuery)%>&amp;page=<%=(abspage + 1)%>&amp;pagesize=<%=pagesize%>"><font size="1"><%=langNext%>&nbsp;&raquo;</font></a>
<%			end if
			i = 0
%>
	</p>

		<table align="center" border="1" width="100%">
		<tr>
<%			for each fld in rec.Fields%>
				<th><%=fld.Name%></th>
<%			next%>
		</tr>

<%
			strFilter = BuildFilter(rec)
			if Len(strFilter) > 0 then call rec.Find(strFilter)
			do while not rec.EOF and i < rec.PageSize and rec.State <> adStateClosed
				if sClass = "oddrow" then sClass = "evenrow" else sClass = "oddrow"
%>
		<tr class="<%=sClass%>" onmouseover="style.backgroundColor='#ffdfbf'" onmouseout="style.backgroundColor=''">
<%				for each fld in rec.Fields%>
					<td valign="top" align="center">
<%					if fld.Type <> adBinary then
						if fld.Value <> "" then Response.Write Replace(fld.Value, "<", "&lt;") else Response.Write "&nbsp;"
					else
						Response.Write "&lt;" & langBinaryData & "&gt;"
					end if
%>
					</td>
<%				next%>
</tr>
<%				rec.MoveNext
				i = i + 1 
			loop
%>

</table>		

	<p align="left">
<%			if abspage > 1 then%>
				<a href="ftquery.asp?query=<%=Server.URLEncode(strQuery)%>&amp;page=<%=(abspage - 1)%>&amp;pagesize=<%=pagesize%>"><font size="1">&laquo;&nbsp;<%=langPrev%></font></a>
<%			end if%>
<%			for i=1 to rec.PageCount
				if i = abspage then%>
					<font size="2">[<%=i%>]</font>&nbsp;
<%				else%>
					<font size="1">&nbsp;[<a href="ftquery.asp?query=<%=Server.URLEncode(strQuery)%>&amp;page=<%=i%>&amp;pagesize=<%=pagesize%>"><%=i%></a>]&nbsp;</font>
<%				end if
			Next
			if abspage < rec.PageCount and abspage > 0 then%>
				<a href="ftquery.asp?query=<%=Server.URLEncode(strQuery)%>&amp;page=<%=(abspage + 1)%>&amp;pagesize=<%=pagesize%>"><font size="1"><%=langNext%>&nbsp;&raquo;</font></a>
<%			end if%>
	</p>
<%		end if
		call DBA_EndNewTable
	end if
%>
<!--END RESULTS FORM-->


<!--BEGIN QUERY FORM-->
<%	
	DBA_BeginNewTable langFreeTypeQuery, langFreeTypeQueryAlt, "90%", ""
	if dba.HasError then DBA_WriteError Replace(dba.LastError, vbCrLf, "<br>")
%>
	<p align="center"><%=langTypeSQL%></p>
	<form action="ftquery.asp" method="post">
	<table align="center" border="0">
		<tr><td>
			<textarea name="query" rows="10" cols="50"><%=strQuery%></textarea>
		</td></tr>
		<tr><td>
			<input type="checkbox" name="transaction" value="-1">&nbsp;<%=langUseTransaction%>
		</td></tr>
		<tr><td>
			<input type="checkbox" name="ignore_errors" value="-1">&nbsp;<%=langIgnoreErrors%>
		</td></tr>
		<tr><td align="center">
			<input class="button" type="submit" value="<%=langRunIt%>" name="submit">
		</td></tr>
	</table>
	</form>

<%
	call DBA_EndNewTable
	set dba = Nothing
%>

<!--END QUERY FORM-->

<!--#include file=scripts/inc_footer.inc-->
</body>
</html>
<%
	Function BuildFilter(ByRef rc)
		dim filter, field, cmp, criteria, fldType, fld
		filter = ""
		field = Request.Form("filter_field").Item
		cmp = Request.Form("filter_cmp").Item
		criteria = Request.Form("filter_criteria").Item
		
		If Len(field) > 0 and Len(criteria) > 0 then
			set fld = new DBAField
			fld.FieldType = rc(field).Type
			fldType = fld.GetSQLTypeName()
			set fld = Nothing
			If fldType = "TEXT" or fldType = "MEMO" Then
				'remove asterics if only at beginning
				If Left(criteria, 1) = "*" and Right(criteria, 1) <> "*" Then criteria = Mid(criteria, 2)
				criteria = "'" & Replace(criteria, "'", "''") & "'"
			ElseIf fldType = "DATETIME" Then
				criteria = "#" & criteria & "#"
			Else
				If cmp = "LIKE" Then cmp = "="
			End If
			filter = field & " " & cmp & " " & criteria
		End If
		
		BuildFilter = filter
	End Function
%>