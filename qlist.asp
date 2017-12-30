<%@ Language=VBScript %>
<!--#include file=scripts\inc_common.asp -->
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<head>
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet">
<title>DBA:Procedures</title>
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
<script language="javascript" type="text/javascript">
<!--
function runSP(sp){
	var params = window.prompt("<%=langEnterQParams%>", "");
	if(params != null && params.length > 0)
		window.location.reload("ftquery.asp?query=" + escape("EXECUTE [" + sp + "] " + params));
}
function deleteSP(name){
	if(confirm('<%=Replace(langDeleteProcedurePrompt, "'", "\'")%>'))
		window.location.href = "qlist.asp?action=delete&name=" + escape(name);
}
//-->
</SCRIPT>
</head>
<body>
<%	call DBA_WriteNavigation%>

<%
	On Error Resume Next
	dim dba, item, sClass, action
	dim EditName, EditBody
	action = CStr(Request("action").Item)
	set dba = new DBAdmin
	dba.Connect Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword)
	
	DBA_BeginNewTable langStoredProceduresList, "", "90%", ""
	if dba.HasError then DBA_WriteError dba.LastError
	Select Case action
		Case "create"
			dba.CreateProcedure Request.Form("pname").Item, Request.Form("pbody").Item
			if dba.HasError then 
				DBA_WriteError dba.LastError
				EditName = Request.Form("pname").Item
				EditBody = Request.Form("pbody").Item
			end if
		Case "edit"
			if dba.Procedures.Exists(Request.QueryString("name").Item) then
				EditName = Request.QueryString("name").Item
				EditBody = dba.Procedures.Item(Request.QueryString("name").Item).Body
				action = "update"
			end if
		Case "update"
			set item = dba.Procedures.Item(Request.Form("origname").Item)
			item.Name = Request.Form("pname").Item
			item.Body = Request.Form("pbody").Item
			action = ""
			if dba.HasError then DBA_WriteError dba.LastError
		Case "delete"
			dba.DeleteProcedure Request.QueryString("name").Item
			action = ""
			if dba.HasError then DBA_WriteError dba.LastError
	End Select

	if Len(action) = 0 then action = "create"
%>

<table align="center" border="0" cellpadding="2" cellspacing="1" width="100%">
	<tr>
		<th><%=langSPName%></th>
		<th><%=langSPCode%></th>
		<th><%=langSPActions%></th>
	</tr>
<%	
	for each item in dba.Procedures.Items
		if sClass = "oddrow" then sClass = "evenrow" else sClass = "oddrow"
%>
	<tr class="<%=sClass%>" onmouseover="style.backgroundColor='#ffdfbf'" onmouseout="style.backgroundColor=''">
		<td valign="top"><b><%=item.Name%></b><br><small><%=langCreated & ":&nbsp;" & item.DateCreated%><br><%=langModified & ":&nbsp;" & item.DateModified%></small></td>
		<td valign="top"><%=HighlightSQL(item.Body)%></td>
		<td align="right" valign="top">
			<a href="javascript:runSP('<%=Replace(item.Name, "'", "\'")%>');"><img src="images/run.gif" border="0" width="16" height="16" alt="<%=langSPExecute%>"></a>&nbsp;
			<a href="qlist.asp?action=edit&amp;name=<%=Server.URLEncode(item.Name)%>"><img src="images/edit.gif" width="16" height="16" border="0" alt="<%=langSPEdit%>"></a>&nbsp;
			<a href="javascript:deleteSP('<%=Server.URLEncode(item.Name)%>');"><img src="images/delete.gif" width="16" height="16" border="0" alt="<%=langSPDelete%>"></a>
		</td>
	</tr>
<%	next%>
</table>

<form action="qlist.asp" method="POST" id="Form1">
<input type="hidden" name="action" value="<%=action%>" id="iAction">
<input type="hidden" name="origname" value="<%=EditName%>" id="iOrigName">
<table align="center" border="0">
	<tr>
		<th align="center"><font size="4"><b><%if action = "update" then Response.Write langUpdateProcedure else Response.Write langCreateProcedure%></b></font></th>
	</tr>
	<tr>
		<th align="center"><%=langCreateProcedureNote%></th>
	</tr>
	<tr><td align="center"><b><%=langProcedureName%></b></td></tr>
	<tr><td align="center"><input type="text" name="pname" id="iPName" value="<%=EditName%>"></td></tr>
	<tr><td align="center"><b><%=langSQLStatement%></b><br>
		<%=langParams1stWay%><br>
		<table align="center"><tr><td><font face="Courier New" color="Green">
		PARAMETERS <em>Param1</em> LONG, <em>Param2</em> TEXT(255);<br>
		SELECT * FROM Table1 WHERE Column1=<em>Param1</em> OR Column2=<em>Param2</em>;
		</font></td></tr></table>
		<%=langParams2ndWay%>
	</td></tr>
	<tr>
		<td align="center"><textarea name="pbody" cols="50" rows="6" id="taPBody"><%=EditBody%></textarea></td>
	</tr>
	<tr>
		<td align="center"><input type="submit" name="submit" value="<%if action = "update" then Response.Write langUpdateProcedure else Response.Write langCreateProcedure%>" class="button" id="Submit1"></td>
	</tr>
</table>
</form>
<%
	call DBA_EndNewTable
	set dba = Nothing
%>
<!--#include file=scripts\inc_footer.inc -->
</body>
</html>
