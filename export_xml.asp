<%@ Language=VBScript %>
<!--#include file=inc_config.asp -->
<!--#include file=inc_protect.asp -->
<!--#include file=inc_functions.asp -->
<%
On Error Resume Next
dim con, rec, s, xml
OpenConnection con
IsError
set rec = Server.CreateObject("ADODB.Recordset")
rec.CursorLocation = adUseClient
rec.Open Request.QueryString("sql").Item, con

set s = Server.CreateObject("ADODB.Stream")
Response.AddHeader "Content-Disposition", "attachment; filename=" & Session.SessionID & "_export.xml"
Response.CharSet  = "UTF-8"
Response.ContentType = "application/octet-stream"
rec.Save s, adPersistXML 
Response.Write s.ReadText
s.Close
set s = nothing

rec.Close
con.Close
set rec = nothing
set con = nothing
%>
