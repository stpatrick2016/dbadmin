<%@ Language=VBScript %>
<!--#include file=scripts\inc_common.asp -->
<%
if Request.Form("action") = "export" then

	On Error Resume Next
	dim dba, rec, s, fld
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
	
	set dba = new DBAdmin
	dba.Connect Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword)
	set rec = dba.RunScript(Request.Form("sql").Item, False, True, null)
	
	Randomize
	Response.AddHeader "Content-Disposition", "attachment; filename=" & Int((Rnd() * 10000000)) & "_export.csv"
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
					s = s & DlmText
					if not IsNull(fld.Value) and Len(fld.Value) > 0 then s = s & Replace(fld.Value, DlmText, DlmText & DlmText)
					s = s & DlmText & DlmColumn
				Case adBinary, adLongVarBinary, adVarBinary
					s = s & DlmColumn
				Case Else
					s = s & fld.Value & DlmColumn
			End Select
		next
		if Len(s) > 0 then s = Left(s, Len(s)-1) & DlmRow
		Response.Write s
		Response.Flush
		
		rec.MoveNext
	loop
	Response.Write DlmRow

	rec.Close
	set rec = nothing
	set dba = Nothing
%>
<%else%>
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<head>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet" type="text/css">
<title>DBA:<%=langCaptionExportCSV%></title>
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
</head>
<body>
<%	
	call DBA_WriteNavigation
	DBA_BeginNewTable langExcelExportAlt, "", "90%", ""
%>
<p align="center"><%=langPleaseDefineExp%></p>
<form action="export_csv.asp" method="POST">
	<input type="hidden" name="sql" value="<%=Replace(Request.QueryString("sql").Item, """", "&quot;")%>">
	<input type="hidden" name="action" value="export">
	<table align="center" cellspacing="1" cellpadding="1">
		<tr>
			<td><%=langColumnDelimiter%></td>
			<td>
				<select name="column">
					<option value="TAB">{<%=langTab%>}</option>
					<option value="SPACE">{<%=langSpace%>}</option>
					<option value=";">;</option>
					<option value=",">,</option>
					<option value="OTHER">{<%=langOther%>} --></option>
				</select>&nbsp; <input type="text" name="other" size="2" maxlength="1">
			</td>
		</tr>
		<tr>
			<td><%=langTextQualifier%></td>
			<td>
				<select name="text">
					<option value='"'>"</option>
					<option value="'">'</option>
					<option value="">{None}</option>
				</select>
			</td>
		</tr>
		<tr>
			<td colspan="2"><input type="checkbox" name="nofields" value="1">&nbsp;<%=langNoFieldNames%></td>
		</tr>
		<tr>
			<td colspan="2" align="center"><input type="submit" name="submit" value="Export" class="button"></td>
		</tr>
	</table>
</form>
<%	call DBA_EndNewTable%>
<!--#include file=scripts\inc_footer.inc -->
</body>
</html>
<%end if%>
