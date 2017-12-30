<%@ Language=VBScript %>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet">
<title>Database Administration</title>
<script LANGUAGE="javascript">
<!--
function showIndexes(){
	var oTable = document.getElementById("tblIndexes");
	if(oTable != null){
		if(oTable.style.display != "none"){
			oTable.style.display = "none";
		}else{
			oTable.style.display = "";
		}
	}
}
//-->
</script>
</head>
<body>
<!--#include file=config.asp -->
<!--#include file=inc_protect.asp -->
<!--#include file=inc_functions.asp -->
<table WIDTH="100%" ALIGN="center">
	<tr>
		<td width="180px" valign="top"><!--#include file=inc_nav.asp --></td>
		<td>
<%
	On Error Resume Next
	dim con, rec, sTableName, sSQL, field, fld,autonumber, sColumns, oIndexes, sError, arIndexes
	pk = ""
	sTableName = Request("table")
	field = Request("field")

	set con = Server.CreateObject("ADODB.Connection")
	con.Open strProvider & Session("DBAdminDatabase") & Session("DBAdminAuth")
	IsError
	set rec = Server.CreateObject("ADODB.Recordset")
	
	'new field
	if Request("submit").Count > 0 then
		if Request.Form("edit") = "1" then
			sSQL = "ALTER TABLE [" & sTableName & "] ALTER COLUMN [" & Request.Form("name") & "] " & Request("type")
		else
			sSQL = "ALTER TABLE [" & sTableName & "] ADD COLUMN [" & Request("name") & "] " & Request("type")
		end if
		if Request("type") = "TEXT" then sSQL = sSQL & "(" & Request("length") & ")"
		if Request.Form("maynull") <> "1" then sSQL = sSQL & " NOT NULL"
		if Request("unique") = "1" and Request("pk") <> "1" then sSQL = sSQL & " UNIQUE"
		if Request("pk") = "1" and Request.Form("edit") <> "1" then sSQL = sSQL & " PRIMARY KEY"
		if Len(Request.Form("default")) > 0 then 
			if Request.Form("type") = "TEXT" or Request.Form("type") = "MEMO" then
				sSQL = sSQL & " DEFAULT """ & Request.Form("default") & """"
			else
				sSQL = sSQL & " DEFAULT " & Request.Form("default")
			end if
		elseif Request.Form("edit") = "1" then sSQL = sSQL & " DEFAULT"
		end if
		con.Execute sSQL, adExecuteNoRecords
		if Err <> 0 then
			Response.Write "<P class=Error align=center>" & Err.Description & "</P>"
'			Response.Write sSQL
		end if
	end if

	set oIndexes = new TableIndexes
	oIndexes.OpenTable sTableName
	
	'# add new Primary Key
	if Request.QueryString("action") = "key" then
		sError = oIndexes.CreateIndex("PrimaryKey", field, True, True)
		if Len(sError) > 0 then
			Response.Write "<P class=Error align=center>" & sError & "</P>"
		end if
	end if
	
	'delete field
	if Request.QueryString("action") = "delete" then
		set rec = con.OpenSchema(adSchemaIndexes, Array(empty,empty,empty,empty, sTableName))
		do while not rec.EOF
			if rec("COLUMN_NAME") = field then
				oIndexes.DeleteIndex rec("INDEX_NAME")
			end if
			rec.MoveNext
		loop
		sSQL = "ALTER TABLE [" & sTableName & "] DROP COLUMN [" & field & "]"
		con.Execute sSQL, adExecuteNoRecords
		if Err <> 0 then
			Response.Write "<P class=Error align=center>Error deleting column: " & Err.Description & "</P>"
		End if
		rec.close
	end if
	
	'delete index
	if Request.QueryString("action") = "delete_index" then
		sError = oIndexes.DeleteIndex(Request.QueryString("index"), field)
		If Len(sError) > 0 then
			Response.Write "<P class=Error align=center>Error deleting index: " & sError & "</P>"
		End if
	end if
	
	'create index
	if Request.Form("create_index").Count > 0 then
		sError = oIndexes.CreateIndex(Request.Form("index_name"), Request.Form("column_name"), False, CBool(Request.Form("unique")))
		if Len(sError) > 0 then
			Response.Write "<P align=center class=error>" & sError & "</P>"
		end if
	end if

	'find the autonumber field
	autonumber = ""
	rec.Open sTableName, con, adOpenForwardOnly, adLockReadOnly, adCmdTable
	if Err = 0 then
		for each fld in rec.Fields
			if fld.type = adInteger and (fld.Attributes and adFldIsNullable) = 0 then
				autonumber = fld.Name
				Exit For
			end if
		next
		rec.Close 
	end if
%>
<p align="center">
<h3 align="center">Table's Indexes</h3>
<p align="center"><a href="javascript:showIndexes()">Show/Hide Indexes table</a></p>
<table align="center" width="100%" border="1" cellpadding="1" cellspacing="1" id="tblIndexes" style="display:none" class=RegularTable>
	<tr>
		<th class=RegularTH>Index Name</th>
		<th class=RegularTH>Column</th>
		<th class=RegularTH>Unique</th>
		<th class=RegularTH>Action</th>
	</tr>
<%
	arIndexes = oIndexes.GetIndexes
	for i=0 to UBound(arIndexes)-1
		if Len(arIndexes(i)(0)) > 0 then
%>
	<tr onmouseover="bgColor='#DDDDDD'" onmouseout="bgColor=''">
		<td class=RegularTD>
<%			if oIndexes.IsPrimaryKey(arIndexes(i)(1)) then%>
				<img src="images/key.gif" border="0" WIDTH="16" HEIGHT="16" alt="Primary Key column">
<%			end if%>
		<%=arIndexes(i)(0)%>
		</td>
		<td class=RegularTD><%=arIndexes(i)(1)%></td>
		<td align="center" class=RegularTD>
			<%if CBool(arIndexes(i)(2)) = True then%>
			<img src="images/check.gif" border="0" WIDTH="16" HEIGHT="16" alt="The index is unique">
			<%end if%>
			&nbsp;
		</td>
		<td align="center" class=RegularTD>
			<a href="structure.asp?table=<%=Server.URLEncode(sTableName)%>&amp;action=delete_index&amp;index=<%=Server.URLEncode(arIndexes(i)(0))%>&field=<%=Server.URLEncode(arIndexes(i)(1))%>"><img src="images/delete.gif" alt="Delete the index" border="0" WIDTH="16" HEIGHT="16"></a>
		</td>
	</tr>
<%		
		end if
		if Err then Exit For
	Next
%>
	<form action="<%=script%>" method=post>
	<input type=hidden name=table value="<%=sTableName%>">
	<tr>
		<td><input type=text name=index_name size=15></td>
		<td><input type=text name=column_name size=15></td>
		<td align=center><input type=checkbox name=unique value="1"></td>
		<td align=right><input type=submit name=create_index value="Create index" class=button></td>
	</TR>
	</form>
</table></p>

<!--Columns definitions-->
<%
	'open schema
	set rec = con.OpenSchema(adSchemaColumns, Array(empty,empty,sTableName))
	if Err then
		Response.Write "<P class=Error align=center>Error opening columns schema: " & Err.Description & "</P>"
	end if
%>
<h1>Table structure for table: <%=sTableName%></h3>
<table WIDTH="100%" ALIGN="center" BORDER="1" CELLSPACING="1" CELLPADDING="1" class=RegularTable>
	<tr>
		<th class=RegularTH>Ordinal</th>
		<th class=RegularTH>Name</th>
		<th class=RegularTH>Data type</th>
		<th class=RegularTH>Nullable</th>
		<th class=RegularTH>Max. length</th>
		<th class=RegularTH>Default Value</th>
		<th class=RegularTH>Description</th>
		<th class=RegularTH>Actions</th>
	</tr>
<%	do while not rec.EOF and Err = 0

		'continue to build create table query
		sColumns = sColumns & rec("ORDINAL_POSITION") & ";[" & rec("COLUMN_NAME") & "] " & GetSQLTypeName(rec("DATA_TYPE"), rec("CHARACTER_MAXIMUM_LENGTH"))
		if rec("DATA_TYPE") = 128 or (rec("DATA_TYPE") = 130 and rec("CHARACTER_MAXIMUM_LENGTH") > 0) then sColumns = sColumns & "(" & rec("CHARACTER_MAXIMUM_LENGTH") & ")"
		if not rec("IS_NULLABLE") then sColumns = sColumns & " NOT NULL"
		if rec("COLUMN_HASDEFAULT") then 
			sColumns = sColumns & " DEFAULT "
			if rec("DATA_TYPE") = 130 then 
				sColumns = sColumns & """" & rec("COLUMN_DEFAULT") & """"
			else
				sColumns = sColumns & rec("COLUMN_DEFAULT")
			end if
		end if
		sColumns = sColumns & vbTab
		
		
		if Request.QueryString("action") = "edit" and field = rec("COLUMN_NAME") then%>
		
		<!--Field editing form-->
	<tr>
		<td class=RegularTD align=center><%=rec("ORDINAL_POSITION")%></td>
		<td class=RegularTD>
			<%if oIndexes.IsPrimaryKey(rec("COLUMN_NAME")) then%>
				<img src="images/key.gif" border="0" WIDTH="16" HEIGHT="16" alt="Primary Key column">
			<%end if%>
			<%=rec("COLUMN_NAME")%>
		</td>
		<form action="structure.asp" method="POST" id="form1" name="form1">
		<input type="hidden" name="edit" value="1">
		<input type="hidden" name="name" value="<%=rec("COLUMN_NAME")%>">
		<input type="hidden" name="table" value="<%=sTableName%>">
			<td class=RegularTD>
				<select name="type">
					<option value="<%=GetSQLTypeName(rec("DATA_TYPE"), rec("CHARACTER_MAXIMUM_LENGTH"))%>"><%=GetTypeName(rec("DATA_TYPE"), rec("CHARACTER_MAXIMUM_LENGTH"))%></option>
					<option value="DATETIME">Date/Time</option>
					<option value="LONG">Long Integer</option>
					<option value="TEXT">Text</option>
					<option value="COUNTER">AutoNumber</option>
					<option value="MEMO">Memo</option>
					<option value="MONEY">Currency</option>
					<option value="BINARY">Binary</option>
					<option value="TINYINT">Byte</option>
					<option value="DECIMAL">Decimal</option>
					<option value="FLOAT">Double</option>
					<option value="INTEGER">Integer</option>
					<option value="REAL">Single</option>
					<option value="BIT">Boolean</option>
					<option value="UNIQUEIDENTIFIER">Replication ID</option>
				</select>
			</td>
			<td class=RegularTD align=center>
				<%if rec("IS_NULLABLE") = True then%>
					<input type=hidden name=maynull value="1" checked>
					True
				<%else%>
					False
				<%end if%>
			</td>
			<td class=RegularTD><input type="text" name="length" size=5 value="<%=rec("CHARACTER_MAXIMUM_LENGTH")%>"></td>
			<td class=RegularTD><input type="text" name="default" size=20 value="<%=rec("COLUMN_DEFAULT")%>"></td>
			<td class=RegularTD align=right colspan=2>
				<input type="submit" name="submit" value="Update" class=button>
				<input type="button" value="Cancel" class=button onclick="javascript:window.location.replace('<%=script%>?table=<%=Server.URLEncode(sTableName)%>');">
			</td>
		</form>
		<!--End of form-->
		<%else%>
		
	<tr onmouseover="bgColor='#DDDDDD'" onmouseout="bgColor=''">
		<td class=RegularTD align=center><%=rec("ORDINAL_POSITION")%></td>
		<td class=RegularTD>
			<%if oIndexes.IsPrimaryKey(rec("COLUMN_NAME")) then%>
				<img src="images/key.gif" border="0" WIDTH="16" HEIGHT="16" alt="Primary Key column">
			<%end if%>
			<%=rec("COLUMN_NAME")%>
		</td>
		<td class=RegularTD>
			<%if rec("COLUMN_NAME") = autonumber then%>
				AutoNumber
			<%else%>
				<%=GetTypeName(rec("DATA_TYPE"), rec("CHARACTER_MAXIMUM_LENGTH"))%>
			<%end if%>
		</td>
		<td align="center" class=RegularTD>
			<%if rec("IS_NULLABLE") then%>
			<img src="images/check.gif" border="0" WIDTH="16" HEIGHT="16">
			<%end if%>
			&nbsp;
		</td>
		<td class=RegularTD><%=rec("CHARACTER_MAXIMUM_LENGTH")%>&nbsp;</td>
		<td class=RegularTD><%=rec("COLUMN_DEFAULT")%>&nbsp;</td>
		<td class=RegularTD><%=rec("DESCRIPTION")%>&nbsp;</td>
		<td align="center" class=RegularTD>
			<a href="structure.asp?table=<%=Server.URLEncode(sTableName)%>&amp;field=<%=Server.URLEncode(rec("COLUMN_NAME"))%>&amp;action=edit"><img src="images/edit.gif" alt="Edit the field" border="0" WIDTH="16" HEIGHT="16"></a>&nbsp;
			<%if oIndexes.IsPrimaryKey(rec("COLUMN_NAME")) then%>
				<a href="structure.asp?table=<%=Server.URLEncode(sTableName)%>&amp;field=<%=Server.URLEncode(rec("COLUMN_NAME"))%>&index=<%=Server.URLEncode(oIndexes.GetPrimaryKeyName)%>&amp;action=delete_index"><img src="images/un_key.gif" alt="Remove Primary Key" border="0" WIDTH="16" HEIGHT="16"></a>&nbsp;
			<%else%>
				<a href="structure.asp?table=<%=Server.URLEncode(sTableName)%>&amp;field=<%=Server.URLEncode(rec("COLUMN_NAME"))%>&amp;action=key"><img src="images/key.gif" alt="Set as Primary Key" border="0" WIDTH="16" HEIGHT="16"></a>&nbsp;
			<%end if%>
			<a href="javascript:deleteColumn('<%=Server.URLEncode(rec("COLUMN_NAME"))%>')"><img src="images/delete.gif" alt="Delete the field" border="0" WIDTH="16" HEIGHT="16"></a>
		</td>
		<%end if%>
	</tr>
<%		rec.MoveNext 
	loop
%>


<!--New field definition-->
<%if Request.QueryString("action") <> "edit" then%>
<form action="structure.asp" method="POST" id="form1" name="form1">
<input type="hidden" name="table" value="<%=sTableName%>">
<tr><TD colspan="8" align=center><B>Add new column</B></TD></TR>
<tr>
	<td class=RegularTD align=center>*<%=UBound(Split(sColumns, vbTab))+1%>*</td>
	<td><input type="text" name="name"></td>
	<td class=RegularTD>
		<select name="type">
			<option value="DATETIME">Date/Time</option>
			<option value="LONG">Long Integer</option>
			<option value="TEXT">Text</option>
			<option value="COUNTER">AutoNumber</option>
			<option value="MEMO">Memo</option>
			<option value="MONEY">Currency</option>
			<option value="BINARY">Binary</option>
			<option value="TINYINT">Byte</option>
			<option value="DECIMAL">Decimal</option>
			<option value="FLOAT">Double</option>
			<option value="INTEGER">Integer</option>
			<option value="REAL">Single</option>
			<option value="BIT">Boolean</option>
			<option value="UNIQUEIDENTIFIER">Replication ID</option>
		</select>
	</td>
	<td class=RegularTD align=center><input type="checkbox" name="maynull" value="1" id="maynull"></td>
	<td class=RegularTD><input type="text" name="length" size=5></td>
	<td class=RegularTD><input type="text" name="default" size=20></td>
	<td class=RegularTD><input type="checkbox" name="unique" id="unique" value="1">Unique field</td>
	<td class=RegularTD align=center><input type="submit" name="submit" value="Add" class=button></td>
</tr>
</form>
<%end if%>
<!--End new field definition-->


</table>
		

<!--CREATE TABLE query-->
<P>
<TABLE align=center class=RegularTable width="75%">
	<TR>
		<TH>CREATE TABLE query for this table<BR>
			<font size=1>Table definition is separated from its indexes<BR>
			Also foreign keys have to be created separately</font>
		</TH>
	</TR>
	<TR><TD style="border:1px inset"><%=BuildTableDef()%></TD></TR>
</TABLE>
</P>

		</td>
	</tr>
</table>
</body>
</html>
<%
	rec.Close
	con.Close
	set rec = nothing
	set con = nothing
	set oIndexes = nothing
%>
<script LANGUAGE="vbscript" RUNAT="Server">
Function GetTypeName(intType, intLength)
	Select Case intType
	Case 3		GetTypeName = "Long Integer"
	Case 7		GetTypeName = "Date/Time"
	Case 11		GetTypeName = "Boolean"
	Case 6		GetTypeName = "Currency"
	Case 128	GetTypeName = "Binary"
	Case 17		GetTypeName = "Byte"
	Case 131	GetTypeName = "Decimal"
	Case 5		GetTypeName = "Double"
	Case 2		GetTypeName = "Integer"
	Case 4		GetTypeName = "Single"
	Case 72		GetTypeName = "Replication ID"
	Case 130
		if intLength = 0 then
			GetTypeName = "Memo"
		else
			GetTypeName = "Text"
		end if
	Case Else	GetTypeName = intType
	End Select
End Function

Function GetSQLTypeName(intType, intLength)
	Select Case intType
	Case 3		GetSQLTypeName = "LONG"
	Case 7		GetSQLTypeName = "DATETIME"
	Case 11		GetSQLTypeName = "BIT"
	Case 6		GetSQLTypeName = "MONEY"
	Case 128	GetSQLTypeName = "BINARY"
	Case 17		GetSQLTypeName = "TINYINT"
	Case 131	GetSQLTypeName = "DECIMAL"
	Case 5		GetSQLTypeName = "FLOAT"
	Case 2		GetSQLTypeName = "INTEGER"
	Case 4		GetSQLTypeName = "REAL"
	Case 72		GetSQLTypeName = "UNIQUEIDENTIFIER"
	Case 130
		if intLength = 0 then
			GetSQLTypeName = "MEMO"
		else
			GetSQLTypeName = "TEXT"
		end if
	Case Else	GetSQLTypeName = "TEXT"
	End Select
End Function

Function BuildTableDef
	dim sTableDef, s, arTemp, arFields, i
	sTableDef = "CREATE TABLE [" & sTableName & "] ("
	arTemp = Split(sColumns, vbTab)
	ReDim arFields(UBound(arTemp))
	for each s in arTemp
		if Len(s) > 0 then 
			i = Left(s, InStr(1, s, ";") - 1)
			arFields(i) = Replace(s, i & ";", "", 1, 1)
		end if
	next
	for each s in arFields
		if Len(s) > 0 then sTableDef = sTableDef & s & ","
	next

	if InStrRev(sTableDef, ",") = Len(sTableDef) then sTableDef = Left(sTableDef, Len(sTableDef)-1)	
	sTableDef = sTableDef & ")"
	if InStrRev(sTableDef, "()") = Len(sTableDef)-1 then sTableDef = Left(sTableDef, Len(sTableDef)-2)	
	'replace LONG with AUTONUMBER
	if Len(autonumber) > 0 then sTableDef = Replace(sTableDef, "[" & autonumber & "] LONG", "[" & autonumber & "] COUNTER", 1, 1, vbTextCompare)
	
	'append Primary Keys query
	sTableDef = sTableDef & ";<BR>"
	if Len(oIndexes.GetPrimaryKeyName) > 0 then
		sTableDef = sTableDef & "CREATE INDEX [" & oIndexes.GetPrimaryKeyName() & "] ON [" & sTableName & "]("
		arTemp = oIndexes.GetPrimaryKeys
		for each i in arTemp
			if Len(i) > 0 then sTableDef = sTableDef & "[" & i & "],"
		Next
		if Right(sTableDef, 1) = "," then sTableDef = Left(sTableDef, Len(sTableDef)-1)
		sTableDef = sTableDef & ") WITH PRIMARY;<BR>"
	end if
	
	'Append other indexes query
	arTemp = oIndexes.GetIndexes
	for i=0 to UBound(arTemp)-1
		if not oIndexes.IsPrimaryKey(arTemp(i)(1)) and not oIndexes.IsForeignKey(arTemp(i)(0)) then
			sTableDef = sTableDef & "CREATE "
			if CBool(arTemp(i)(2)) = true then sTableDef = sTableDef & "UNIQUE "
			sTableDef = sTableDef & "INDEX [" & arTemp(i)(0) & "] ON [" & sTableName & "]([" & arTemp(i)(1) & "]);<BR>"
		end if
	Next
	
	BuildTableDef = HighlightSQL(sTableDef)
End Function

</script>

<script LANGUAGE="javascript">
<!--
function deleteColumn(name){
	if(confirm("Are you sure you want to delete column " + name + " and all indexes to it?")){
		document.location.replace("structure.asp?table=<%=Server.URLEncode(sTableName)%>&field=" + name + "&action=delete");
	}
}
//-->
</script>
