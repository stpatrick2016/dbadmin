<%@ Language=VBScript %>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet">
<title>Database Administration</title>
</head>
<body>
<!--#include file=config.asp -->
<!--#include file=inc_protect.asp -->
<!--#include file=inc_functions.asp -->

<table WIDTH="100%" ALIGN="center">
	<tr>
		<td width="180px" valign="top"><!--#include file=inc_nav.asp --></td>
		<td>
		
<h1 align="center">Relationships</h1>
<%
	On Error Resume Next
	dim con, rec, script, sSQL
	script = Request.ServerVariables("SCRIPT_NAME")

	set con = Server.CreateObject("ADODB.Connection")
	con.Open strProvider & Session("DBAdminDatabase") & Session("DBAdminAuth")
	IsError
	
	'create new foreign key
	if Request.Form("submit").Count > 0 then
		sSQL =	"ALTER TABLE [" & Replace(Request.Form("fk_table"), "'", "''") & "] ADD CONSTRAINT [" &_
				Replace(Request.Form("fk_name"), "'", "''") & "] FOREIGN KEY ([" & Replace(Request.Form("fk_column"), "'", "''") &_
				"]) REFERENCES [" & Replace(Replace(Request.Form("primary_table"), ".", "] (["), "'", "''") &_
				"])"
		if Request.Form("update") <> "" then sSQL = sSQL & " ON UPDATE " & Request.Form("update")
		if Request.Form("delete") <> "" then sSQL = sSQL & " ON DELETE " & Request.Form("delete")
		con.Execute sSQL, adExecuteNoRecords
		if Err then
			Response.Write "<P align=""center"" class=""error"">" & Err.Description & "</P>"
		end if
	end if
	
	'delete foreign key
	if Request.QueryString("action") = "delete" then
		sSQL =	"ALTER TABLE [" & Replace(Request.QueryString("fk_table"), "'", "''") & "] DROP CONSTRAINT [" &_
				Replace(Request.QueryString("fk_name"), "'", "''") & "]"
		con.Execute sSQL, adExecuteNoRecords
		if Err then
			Response.Write "<P align=""center"" class=""error"">" & Err.Description & "</P>"
		end if
	end if
	
	set rec = con.OpenSchema(adSchemaForeignKeys)
	
%>	
<p align="center">
Each of relationships described also in more readable form.<br>
Note that only relationships that have any meaning are displayed
</p>


<table align="center" width="100%" border="1" class="RegularTable">
<%	
		do while not rec.EOF and Err=0
			if Len(rec("PK_NAME")) > 0 then
%>
	<tr>
		<th>Primary Index</th>
		<th>Primary Table</th>
		<th>Primary Column</th>
		<th>Foreign Index</th>
		<th>Foreign Table</th>
		<th>Foreign Column</th>
		<th>On Update</th>
		<th>On Delete</th>
	</tr>
	<tr onmouseover="bgColor='#DDDDDD'" onmouseout="bgColor=''">
		<td><%=rec("PK_NAME")%></td>
		<td><%=rec("PK_TABLE_NAME")%></td>
		<td><%=rec("PK_COLUMN_NAME")%></td>
		<td><%=rec("FK_NAME")%></td>
		<td><%=rec("FK_TABLE_NAME")%></td>
		<td><%=rec("FK_COLUMN_NAME")%></td>
		<td><%=rec("UPDATE_RULE")%></td>
		<td><%=rec("DELETE_RULE")%></td>
	</tr>

<%	'Readable form%>
	<tr>
		<td valign="top">
			<b>Description:</b><br>
			<a href="<%=script%>?action=delete&amp;fk_name=<%=Server.URLEncode(rec("FK_NAME"))%>&amp;fk_table=<%=Server.URLEncode(rec("FK_TABLE_NAME"))%>"><img src="images/delete.gif" alt="Delete this relationship" border="0" WIDTH="16" HEIGHT="16"></a>
		</td>
		<td colspan="7">
		<ul>
		<%if rec("UPDATE_RULE") <> "NO ACTION" then%>
			<li>If field <b><i><%=rec("PK_COLUMN_NAME")%></i></b> changed in <b><i><%=rec("PK_TABLE_NAME")%></i></b>, 
			then field <b><i><%=rec("FK_COLUMN_NAME")%></i></b> in <b><i><%=rec("FK_TABLE_NAME")%></i></b> will be 
			<%if rec("UPDATE_RULE") = "CASCADE" then%>
				changed also.
			<%elseif rec("UPDATE_RULE") = "SET NULL" then%>
				set to null.
			<%else%>
				set to default value.
			<%end if%>
			</li>
		<%end if%>

		<%if rec("DELETE_RULE") <> "NO ACTION" then%>
			<li>If record with <b><i><%=rec("PK_COLUMN_NAME")%></i></b> deleted from <b><i><%=rec("PK_TABLE_NAME")%></i></b>, 
			<%if rec("DELETE_RULE") = "CASCADE" then%>
				then all records with same <b><i><%=rec("FK_COLUMN_NAME")%></i></b> will be deleted from <b><i><%=rec("FK_TABLE_NAME")%>.</i></b>
			<%elseif rec("DELETE_RULE") = "SET NULL" then%>
				then equal to it field <b><i><%=rec("FK_COLUMN_NAME")%></i></b> in <b><i><%=rec("FK_TABLE_NAME")%></i></b> will be 
				set to null.
			<%else%>
				then equal to it field <b><i><%=rec("FK_COLUMN_NAME")%></i></b> in <b><i><%=rec("FK_TABLE_NAME")%></i></b> will be 
				set to default value.
			<%end if%>
			</li>
		<%end if%>
		</ul></td>
	</tr>
	<tr>
		<td valign=top><B>SQL</B></td>
		<td colspan="7">
		<%
		sSQL =	"ALTER TABLE [" & Replace(rec("FK_TABLE_NAME"), "'", "''") & "] ADD CONSTRAINT [" &_
				Replace(rec("FK_NAME"), "'", "''") & "] FOREIGN KEY ([" & Replace(rec("FK_COLUMN_NAME"), "'", "''") &_
				"]) REFERENCES [" & Replace(rec("PK_TABLE_NAME"), "'", "''") & "] ([" & Replace(rec("PK_COLUMN_NAME"), "'", "''") &_
				"])"
		if rec("UPDATE_RULE") <> "NO ACTION" then sSQL = sSQL & " ON UPDATE " & rec("UPDATE_RULE")
		if rec("DELETE_RULE") <> "NO ACTION" then sSQL = sSQL & " ON DELETE " & rec("DELETE_RULE")
		%>
		<%=HighlightSQL(sSQL)%></td>
	</tr>
	<tr><td colspan="8" bgcolor="white"><hr width="75%"></td></tr>
<%		
			end if
			rec.MoveNext
		loop
		rec.close
%>
</table>


<br><br>
<h3 align="center">Create new relationship</h3>

<form action="<%=script%>" method="POST">
<table border="0" align="center">
	<tr>
		<th>Foreign index name</th>
		<th>Primary table</th>
		<th>Foreign(child) table</th>
		<th>Foreign column</th>
	</tr>
	<tr>
		<td><input type="text" name="fk_name" id="FKName" onchange="bCustomName=true;"></td>
		<td>
			<%set rec = con.OpenSchema(adSchemaPrimaryKeys)%>
			<select name="primary_table" size="1" id="PrimarySelect" onchange="javascript:onColumnChange();">
			<%do while not rec.EOF
				if rec("TABLE_NAME") <> "MSysAccessObjects" and rec("TABLE_NAME") <> "MSysIMEXColumns" and rec("TABLE_NAME") <> "MSysIMEXSpecs" then
			%>
				<option value="<%=rec("TABLE_NAME") & "." & rec("COLUMN_NAME")%>"><%=rec("TABLE_NAME")%></option>
			<%	end if
				rec.MoveNext
			  loop%>
			</select>
			<%rec.Close%>
		</td>
		<td>
			<%
				dim sTables, s1, sLastTable, ar1, i, ar2, s
				set rec = con.OpenSchema(adSchemaColumns)
				s1 = ""
				sLastTable = ""
				sTables = ""
				do while not rec.EOF
					if rec("TABLE_NAME") <> "MSysAccessObjects" and rec("TABLE_NAME") <> "MSysIMEXColumns" and rec("TABLE_NAME") <> "MSysIMEXSpecs" then
						if sLastTable <> rec("TABLE_NAME") then
							sLastTable = rec("TABLE_NAME")
							if Right(sTables, 1) = "." then sTables = Left(sTables, Len(sTables)-1)
							sTables = sTables & "!" & sLastTable & "."
						end if
						sTables = sTables & rec("COLUMN_NAME") & "."
					end if
					rec.MoveNext
				loop
				if Left(sTables, 1) = "!" then sTables = Right(sTables, Len(sTables)-1)
				if Right(sTables, 1) = "." or Right(sTables, 1) = "!" then sTables = Left(sTables, Len(sTables)-1)
				rec.close
				ar1 = Split(sTables, "!")
			%>
			<script language="Javascript">
			var bCustomName = false;
			function onTableChange(){
				var arColumns;
				<%
					Response.Write "arColumns = new Array("
					s = ""
					for each s1 in ar1
						Response.Write s & "new Array("
						ar2 = Split(s1, ".")
						s = ""
						for i=1 to UBound(ar2)
							Response.Write s & "'" & Replace(ar2(i), "'", "\'") & "'"
							s = ","
						next
						response.Write ")"
					next
					Response.Write ");" & vbCrLf
				%>
				
				var oTableSelect = document.getElementById("TableSelect"),
					oColumnSelect = document.getElementById("ColumnSelect");
				var i, oOption;
				if(oTableSelect != null && oColumnSelect != null){
					oColumnSelect.options.length = 0;
					for(i=0; i<arColumns[oTableSelect.selectedIndex].length; i++){
						oOption = document.createElement("OPTION");
						oColumnSelect.appendChild(oOption);
						oOption.value = arColumns[oTableSelect.selectedIndex][i];
						oOption.text = arColumns[oTableSelect.selectedIndex][i];
					}
					var oFKName = document.getElementById("FKName"),
						oPrimarySelect = document.getElementById("PrimarySelect");
					if(!bCustomName && oFKName != null && oPrimarySelect != null){
						oFKName.value = oPrimarySelect[oPrimarySelect.selectedIndex].value.replace(/\./ig, "") + oTableSelect[oTableSelect.selectedIndex].value + oColumnSelect[oColumnSelect.selectedIndex].value;
					}
				}
			}
			function onColumnChange(){
				var oTableSelect = document.getElementById("TableSelect"),
					oColumnSelect = document.getElementById("ColumnSelect"),
					oFKName = document.getElementById("FKName"),
					oPrimarySelect = document.getElementById("PrimarySelect");
				if(!bCustomName && oFKName != null && oPrimarySelect != null && oTableSelect != null && oColumnSelect != null){
					oFKName.value = oPrimarySelect[oPrimarySelect.selectedIndex].value.replace(/\./ig, "") + oTableSelect[oTableSelect.selectedIndex].value + oColumnSelect[oColumnSelect.selectedIndex].value;
				}
			}
			</script>
			<select name="fk_table" id="TableSelect" onchange="javascript:onTableChange();">
			<%for each s1 in ar1%>
				<option value="<%=Split(s1, ".")(0)%>"><%=Split(s1, ".")(0)%></option>
			<%next%>
			</select>
		</td>
		<td>
			<select name="fk_column" id="ColumnSelect" style="width:200px" onchange="javascript:onColumnChange();">
			<%
				ar2 = Split(ar1(0), ".")
				for i=1 to UBound(ar2)
			%>
				<option value="<%=ar2(i)%>"><%=ar2(i)%></option>
			<%	next%>
			</select>
		</td>
	</tr>
	<tr>
		<td colspan="2" align="center"><b>On update:&nbsp;</b>
			<select name="update">
				<option value>No Action</option>
				<option value="CASCADE">Update</option>
				<option value="SET NULL">Set to Null</option>
			</select>
		</td>
		<td colspan="2" align="center"><b>On delete:&nbsp;</b>
			<select name="delete">
				<option value>No Action</option>
				<option value="CASCADE">Delete</option>
				<option value="SET NULL">Set to Null</option>
			</select>
		</td>
	</tr>
</table>
<p align="center"><input type="submit" name="submit" value="Create Relationship" class="button"></p>
</form>



<%
	con.Close
	set rec = nothing
	set con = nothing
%>		
		</td>
	</tr>
</table>


</body>
</html>
