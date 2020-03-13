<%@ Language=VBScript %>
<!--#include file=scripts\inc_common.asp -->
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<head>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet" type="text/css">
<title>DBA:<%=langCaptionViews%></title>
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
<script type="text/javascript" language="javascript">
function deleteView(name){
	var text = '<%=Replace(langSureToDeleteView, "'", "\'")%>';
	text = text.replace('$name', name);
	if(confirm(text))
		window.location.href = "vlist.asp?action=delete&name=" + escape(name);
}
</script>
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
	
	DBA_BeginNewTable langViews, "", "90%", ""
	if dba.HasError then DBA_WriteError dba.LastError
	Select Case action
		Case "create"
			dba.CreateView Request.Form("vname").Item, Request.Form("vbody").Item
			action = ""
			if dba.HasError then 
				DBA_WriteError dba.LastError
				EditName = Request.Form("vname").Item
				EditBody = Request.Form("vbody").Item
			end if
		Case "edit"
			if dba.Views.Exists(Request.QueryString("name").Item) then
				EditName = Request.QueryString("name").Item
				EditBody = dba.Views.Item(Request.QueryString("name").Item).Body
				action = "update"
			end if
		Case "update"
			set item = dba.Views.Item(Request.Form("origname").Item)
			item.Name = Request.Form("vname").Item
			item.Body = Request.Form("vbody").Item
			action = ""
			if dba.HasError then DBA_WriteError dba.LastError
		Case "delete"
			dba.DeleteView Request.QueryString("name").Item
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
	for each item in dba.Views.Items
		if sClass = "oddrow" then sClass = "evenrow" else sClass = "oddrow"
%>
	<tr class="<%=sClass%>" onmouseover="style.backgroundColor='#ffdfbf'" onmouseout="style.backgroundColor=''">
		<td valign="top"><b><%=item.Name%></b><br><small><%=langCreated & ":&nbsp;" & item.DateCreated%><br><%=langModified & ":&nbsp;" & item.DateModified%></small></td>
		<td valign="top"><%=HighlightSQL(item.Body)%></td>
		<td align="right" valign="top">
			<a href="ftquery.asp?query=<%=Server.URLEncode("SELECT * FROM [" & item.Name & "]")%>"><img src="images/run.gif" border="0" width="16" height="16" alt="<%=langRunViewAlt%>"></a>&nbsp;
			<a href="vlist.asp?action=edit&amp;name=<%=Server.URLEncode(item.Name)%>"><img src="images/edit.gif" width="16" height="16" border="0" alt="<%=langEditView%>"></a>&nbsp;
			<a href="javascript:deleteView('<%=Server.URLEncode(item.Name)%>');"><img src="images/delete.gif" width="16" height="16" border="0" alt="<%=langDeleteViewAlt%>"></a>
		</td>
	</tr>
<%	next%>
</table>

<form action="vlist.asp" method="post">
<input type="hidden" name="action" value="<%=action%>" id="iAction">
<input type="hidden" name="origname" value="<%=EditName%>" id="iOrigName">
<table align="center" border="0">
	<tr>
		<th align="center" colspan="2"><font size="4"><b><%if action = "update" then Response.Write langUpdateView else Response.Write langCreateNewView%></b></font></th>
	</tr>
	<tr>
		<td><b><%=langViewName%></b></td>
		<td><input type="text" name="vname" id="iName" value="<%=EditName%>"></td>
	</tr>
	<tr><td colspan="2" align="center"><b><%=langSQLStatement%></b></td></tr>
	<tr><td colspan="2" align="center"><textarea name="vbody" cols="50" rows="6" id="taBody"><%=EditBody%></textarea></td></tr>
	<tr><td align="center" colspan="2">
		<input type="submit" name="submit" value="<%if action = "update" then Response.Write langUpdateView else Response.Write langCreateNewView%>" class="button" id="Submit1">
	</td></tr>
</table>	
</form>
<%
	call DBA_EndNewTable
	set dba = Nothing
%>
<!--#include file=scripts\inc_footer.inc -->
</body>
</html>
