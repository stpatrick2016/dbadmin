<%@ Language=VBScript %>
<!--#include file=scripts/inc_common.asp -->
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>


<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" href="default.css" type="text/css">
<title><%=langLookup%></title>
<script type="text/javascript" language="javascript">
function chooseValue(val){
	window.opener.document.getElementById('<%=Request("id").Item%>').value = val;
	window.close();
}
</script>
</head>
<body>

<%
	If not DBAE_DEBUG Then On Error Resume Next
	dim table, field, backId, dba, pagesize, page, rec, f, i, sClass
	table = Request("table").Item
	field = Request("field").Item
	backId = Request("id").Item
	pagesize = 10
	page = 1
	if IsNumeric(Request("pagesize").Item) then pagesize = CInt(Request("pagesize").Item)
	if IsNumeric(Request("page").Item) then page = CInt(Request("page").Item)
	if pagesize < 1 then pagesize = StpProfile.GetProfileNumber("settings", "page_size", 10)
	if page < 1 then page = 1
	
	If Len(table) = 0 or Len(field) = 0 Then
		Response.Write langNoLookupValues
	Else
		set dba = new DBAdmin
		dba.Connect Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword)
		set rec = dba.Tables.Item(table).GetRawData(pagesize, "", True)
		if rec.RecordCount = 0 Then
			Response.Write langNoLookupValues
		Else
			rec.CacheSize = pagesize
			rec.PageSize = pagesize
			rec.AbsolutePage = page
			call ShowValues()
		End If
		rec.close
		set rec = Nothing
		set dba = Nothing
	End If
%>

<%
Sub ShowValues
	Response.Write	"<p align=""center"">" & langLookupAlt & "</p>" &_
					getPagingControl(rec.RecordCount, page, pagesize, "&amp;table=" & Server.URLEncode(table) & "&amp;field=" & Server.URLEncode(field) & "&amp;id=" & backId) &_
					"<table align=""center"" width=""100%"" border=""0""><tr>"
	for each f in rec.Fields
		Response.Write "<th>" & f.Name & "</th>"
	next
	Response.Write "</tr>"
	
	i = 0
	do while not rec.EOF and i < pagesize
		if sClass = "oddrow" then sClass = "evenrow" else sClass = "oddrow"
		Response.Write	"<tr class=""" & sClass & """ onmouseover=""style.backgroundColor='#ffdfbf'"" onmouseout=""style.backgroundColor=''"">" &_
						"<td><a href=""javascript:chooseValue('" & rec(field).Value & "')"">" & rec(field).Value & "</a></td>"
		for each f in rec.Fields
			If f.Name <> field Then Response.Write "<td>" & Left(f.Value & " ", 20) & "</td>"
		Next
		Response.Write "</tr>" & vbCrLf
		call rec.MoveNext()
		i = i + 1
	loop
	Response.Write	"</table>" &_
					getPagingControl(rec.RecordCount, page, pagesize, "&amp;table=" & Server.URLEncode(table) & "&amp;field=" & Server.URLEncode(field) & "&amp;id=" & backId)
End Sub
%>

<p align="center"><a href="javascript:window.close();"><%=langClose%></a></p>
</body>
</html>

