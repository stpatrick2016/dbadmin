<%@ Language=VBScript %>
<!--#include file=scripts\inc_common.asp -->
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<head>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet" type="text/css">
<title>DBA:<%=langRelations%></title>
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
<script type="text/javascript" language="javascript">
//<!--
var arrTables = new Array();
var dbtbl = null;
var bCustomName = false;

function DBField(name, isPrimary){
	this.name = name;
	this.IsPrimary = isPrimary;
}
function DBTable(){
	this.name = "";
	this.fields = new Array();
	
	this.addField = function(name, isPrimary){this.fields.push(new DBField(name, isPrimary));};
}

function loadTables(){
	var fk, pk;
	fk = document.getElementById('FKTable');
	pk = document.getElementById('PKTable');
	var i, j;
	for(i = 0; i<arrTables.length; i++){
		fk.options[fk.options.length] = new Option(arrTables[i].name, arrTables[i].name, false, false);

		for(j=0; j<arrTables[i].fields.length; j++){
			if(arrTables[i].fields[j].IsPrimary){
				pk.options[pk.options.length] = new Option(arrTables[i].name, arrTables[i].name, false, false);
				break;
			}
		}
	}
}
function onPKTableChange(){
	var tbl, fld;
	tbl = document.getElementById('PKTable');
	fld = document.getElementById('PKField');
	fld.options.length = 0;
	for(var i=0; i<arrTables.length; i++){
		if(arrTables[i].name == tbl.options[tbl.selectedIndex].value){
			for(var j=0; j<arrTables[i].fields.length; j++){
				if(arrTables[i].fields[j].IsPrimary){
					fld.options[fld.options.length] = new Option(arrTables[i].fields[j].name, arrTables[i].fields[j].name, false, false);
				}
			}
			break;
		}
	}
	buildFKName();
}
function onFKTableChange(){
	var tbl, fld;
	tbl = document.getElementById('FKTable');
	fld = document.getElementById('FKField');
	fld.options.length = 0;
	for(var i=0; i<arrTables.length; i++){
		if(arrTables[i].name == tbl.options[tbl.selectedIndex].value){
			for(var j=0; j<arrTables[i].fields.length; j++){
				fld.options[fld.options.length] = new Option(arrTables[i].fields[j].name, arrTables[i].fields[j].name, false, false);
			}
			break;
		}
	}
	buildFKName();
}
function buildFKName(){
	if(bCustomName) return;
	
	var pkTbl, pkFld, fkTbl, fkFld, oName;
	fkTbl = document.getElementById('FKTable');
	fkFld = document.getElementById('FKField');
	pkTbl = document.getElementById('PKTable');
	pkFld = document.getElementById('PKField');
	oName = document.getElementById('FKName');
	
	oName.value =	"FK_" + 
					pkTbl.options[pkTbl.selectedIndex].value + pkFld.options[pkFld.selectedIndex].value +
					"_" +
					fkTbl.options[fkTbl.selectedIndex].value + fkFld.options[fkFld.selectedIndex].value;
}
//-->
</script>
</head>
<body>
<%	call DBA_WriteNavigation%>

<%
	On Error Resume Next
	dim dba, key, dic, s, fld, tbl
	set dba = new DBAdmin
	dba.Connect Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword)
	If not dba.IsOpen then Response.Redirect "database.asp"
	
	DBA_BeginNewTable langRelations, langRelationsNote, "90%", ""
	if dba.HasError then DBA_WriteError dba.LastError
	if Request.Form("submit").Count > 0 then 
		dba.CreateRelation Request.Form("fk_name"), Request.Form("pk_table"), Request.Form("pk_field"), Request.Form("fk_table"), Request.Form("fk_field"), Request.Form("onupdate"), Request.Form("ondelete")
		if dba.HasError then DBA_WriteError dba.LastError
	end if
	
	if Request.QueryString("action") = "delete" then
		dba.DeleteRelation Request.QueryString("fk_name"), Request.QueryString("fk_table")
	end if

	
	'write out javascript
	set dic = dba.Tables
	Response.Write "<sc" & "ript language=""javascript"" type=""text/javascript"">" & vbCrLf
	for each tbl in dic.Items
		Response.Write	"dbtbl = new DBTable;" & vbCrLf &_
						"dbtbl.name = '" & Replace(tbl.Name, "'", "\'") & "';" & vbCrLf
		for each fld in tbl.Fields.Items
			Response.Write	"dbtbl.addField('" & Replace(fld.Name, "'", "\'") & "', " & CInt(fld.IsPrimaryKey) & ");" & vbCrLf
		next
		Response.Write	"arrTables.push(dbtbl);" & vbCrLf & vbCrLf
	next
	Response.Write "</sc" & "ript>" & vbCrLf
	
	set dic = dba.Relations
%>

<table align="center" border="0" cellpadding="3" cellspacing="1" width="90%">
<%	
	for each key in dic.Keys
%>
	<tr>
		<th colspan="7" align="center"><%=langForeignIndex%>:&nbsp;<i><%=dic.Item(key).Name%></i></th>
	</tr>
	<tr>
		<th><%=langPrimaryIndex%></th>
		<th><%=langPrimaryTable%></th>
		<th><%=langPrimaryColumn%></th>
		<th><%=langForeignTable%></th>
		<th><%=langForeignColumn%></th>
		<th><%=langOnUpdate%></th>
		<th><%=langOnDelete%></th>
	</tr>
	<tr class="evenrow">
		<td><%=dic.Item(key).PrimaryIndex%></td>
		<td><%=dic.Item(key).PrimaryTable%></td>
		<td><%=dic.Item(key).PrimaryField%></td>
		<td><%=dic.Item(key).ForeignTable%></td>
		<td><%=dic.Item(key).ForeignField%></td>
		<td><%=dic.Item(key).OnUpdate%></td>
		<td><%=dic.Item(key).OnDelete%></td>
	</tr>

<%	'Readable form%>
	<tr class="evenrow">
		<td valign="top">
			<b><%=langDescription%></b><br>
			<a href="relations.asp?action=delete&amp;fk_name=<%=Server.URLEncode(dic.Item(key).Name)%>&amp;fk_table=<%=Server.URLEncode(dic.Item(key).ForeignTable)%>"><img src="images/delete.gif" alt="<%=langDeleteRelationship%>" border="0" width="16" height="16"></a>
		</td>
		<td colspan="6">
		<ul>
		<%
		if dic.Item(key).OnUpdate <> "NO ACTION" then
			s = Replace(langIfFieldChanged, "$PK_COLUMN_NAME", dic.Item(key).PrimaryField)
			s = Replace(s, "$PK_TABLE_NAME",dic.Item(key).PrimaryTable)
			s = Replace(s, "$FK_COLUMN_NAME", dic.Item(key).ForeignField)
			s = Replace(s, "$FK_TABLE_NAME",dic.Item(key).ForeignTable)
			if dic.Item(key).OnUpdate = "CASCADE" then
				s = s & langChangedAlso
			elseif dic.Item(key).OnUpdate = "SET NULL" then
				s = s & langSetToNull
			else
				s = s & langSetToDefault
			end if
		%>
			<li><%=s%></li>
		<%end if%>

		<%
		if dic.Item(key).OnDelete <> "NO ACTION" then
			s = Replace(langIfFieldDeleted, "$PK_COLUMN_NAME", dic.Item(key).PrimaryField)
			s = Replace(s, "$PK_TABLE_NAME",dic.Item(key).PrimaryTable)
			s = Replace(s, "$FK_COLUMN_NAME", dic.Item(key).ForeignField)
			s = Replace(s, "$FK_TABLE_NAME",dic.Item(key).ForeignTable)
			if dic.Item(key).OnDelete = "CASCADE" then
				s = s & langWillBeDeleted
			elseif dic.Item(key).OnDelete = "SET NULL" then
				s = s & langSetToNull
			else
				s = s & langSetToDefault
			end if
		%>
			<li><%=s%></li>
		<%end if%>
		</ul></td>
	</tr>
	<tr class="evenrow">
		<td valign=top><b>SQL</b></td>
		<td colspan="6"><%=HighlightSQL(dic.Item(key).SQL)%></td>
	</tr>
	<tr><td colspan="7"><hr width="75%"></td></tr>
<%
	next
%>
</table>

<%
	call DBA_EndNewTable
	DBA_BeginNewTable langCreateRelationship, "", "90%", ""
%>
<form action="relations.asp" method="post">
<table align="center" border="0">
	<tr>
		<th><%=langForeignIndexName%></th>
		<th><%=langPrimaryTable%></th>
		<th><%=langPrimaryColumn%></th>
		<th><%=langForeignTable%></th>
		<th><%=langForeignColumn%></th>
	</tr>
	<tr>
		<td><input type="text" name="fk_name" id="FKName" onchange="bCustomName = true;"></td>
		<td><select name="pk_table" id="PKTable" onchange="javascript:onPKTableChange();"><option value=""></option></select></td>
		<td><select name="pk_field" id="PKField" onchange="javascript:buildFKName();"><option value=""></option></select></td>
		<td><select name="fk_table" id="FKTable" onchange="javascript:onFKTableChange();"><option value=""></option></select></td>
		<td><select name="fk_field" id="FKField" onchange="javascript:buildFKName();"><option value=""></option></select></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td align="center"><b><%=langOnUpdate%></b>
			<select name="onupdate">
				<option value="NO ACTION"><%=langNoAction%></option>
				<option value="CASCADE"><%=langUpdate%></option>
			</select>
		</td>
		<td>&nbsp;</td>
		<td align="center"><b><%=langOnDelete%></b>
			<select name="ondelete">
				<option value="NO ACTION"><%=langNoAction%></option>
				<option value="CASCADE"><%=langDelete%></option>
			</select>
		</td>
		<td>&nbsp;</td>
	</tr>
	<tr><td colspan="5" align="center">
		<input type="submit" name="submit" value="<%=langCreateRelationship%>" class="button">
	</td></tr>
</table>
</form>
<%
	call DBA_EndNewTable
	set dba = Nothing
%>
<script type="text/javascript" language="javascript">loadTables();</script>
<!--#include file=scripts\inc_footer.inc -->
</body>
</html>
