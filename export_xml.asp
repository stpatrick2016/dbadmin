<%@ Language=VBScript %>
<!--#include file=scripts\inc_common.asp -->
<%
'	On Error Resume Next
	dim dba, rec, s, xml
	set dba = new DBAdmin
	dba.Connect Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword)
	set rec = dba.RunScript(Request("sql").Item, False, True, null)

	set s = Server.CreateObject("ADODB.Stream")
	rec.Save s, adPersistXML
	Randomize
	Response.AddHeader "Content-Disposition", "attachment; filename=" & Int((Rnd() * 10000000)) & "_export.xml"
	Response.ContentType = "application/octet-stream"
	Response.CharSet  = "UTF-8"
	Response.Write s.ReadText
	s.Close

	set s = nothing

	rec.Close
	set rec = nothing
	set dba = Nothing
%>
