<%@ Language=VBScript %>
<!--#include file=scripts\inc_common.asp -->
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
<title>DBA:<%=langCaptionImportDB%></title>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta name="GENERATOR" content="Microsoft Visual Studio.NET 7.0">
<link href="default.css" rel="stylesheet" type="text/css">
<script type="text/javascript" language="javascript" src="scripts/common.js" defer></script>
<script type="text/javascript" language="javascript">
	var win;
	function browseDB(){
		win = window.open("browse.asp", "browse", "innerHeight=400,height=400,innerWidth=300,width=300,status=no,resizable=no,menubar=no,toolbar=no,center=yes,scrollbars=yes", false);
	}
</script>
</head>
<body>
<%	call DBA_WriteNavigation%>

<%
	dim dba, action
	
	action = lcase(Request("action").Item)
	set dba = new DBAdmin
	call dba.Connect(Session(DBA_cfgSessionDBPathName), Session(DBA_cfgSessionDBPassword))
	if dba.HasError then DBA_WriteError dba.LastError
	
	if action = "show" and Len(Request.Form("e_path").Item) > 0 then
		call ShowExternalTables()
	ElseIf action = "import" then
		call DoImport()
	else
		call AskForDatabase()
	end if
	
	set dba = Nothing
%>

<!--#include file=scripts\inc_footer.inc -->
</body>
</html>

<%
	Sub AskForDatabase
		call DBA_BeginNewTable(langImportDatabase, langImportDatabaseNote, "90%", "")
%>
		<p><%=langImportDatabaseWelcome%></p>
		<form action="import_db.asp" method="post">
		<input type="hidden" name="action" value="show">
		<table align="center" border="0">
			<tr>
				<td><%=langPathToExternalDatabase%></td>
				<td>
					<input type="text" name="e_path" size="30" id="iPath" value="<%=Request("e_path")%>">&nbsp;
					<input type="button" value="Browse" class="button" onclick="javascript:browseDB();">
				</td>
			</tr>
			<tr>
				<td><%=langExternalDBPassword%></td>
				<td><input type="password" name="e_pwd"></td>
			</tr>
			<tr>
				<td colspan="2" align="center"><hr width="75%"></td>
			</tr>
			<tr>
				<td colspan="2" align="center">
					<input type="submit" value="<%=langNext%> >>" class="button">
				</td>
			</tr>
		</table>
		</form>
		
<%
		call DBA_EndNewTable
	End Sub
%>

<%
	Sub ShowExternalTables
		dim extDba, item
		set extDba = new DBAdmin
		call DBA_BeginNewTable(langImportDatabase, "", "90%", "")
		call extDba.Connect(Request("e_path").Item, Request("e_pwd").Item)
		if extDba.HasError then
			call DBA_WriteError(extDba.LastError)
			set extDba = Nothing
			Exit Sub
		End If
%>
		<p><%=langSelectExternalTables%></p>
		<form action="import_db.asp" method="post" id="Form1">
		<input type="hidden" name="action" value="import">
		<input type="hidden" name="e_path" value="<%=Request.Form("e_path")%>">
		<input type="hidden" name="e_pwd" value="<%=Request.Form("e_pwd")%>">
		<table align="center" border="0" cellspacing="1" cellpadding="10">
			<tr class="evenrow">
				<td><b><%=langTablesList%></b><br><select multiple name="table" size="10">
<%		for each item in extDba.Tables.Items%>
					<option value="<%=item.Name%>"><%=item.Name%></option>
<%		next%>
				</select></td>
				
				<td><b><%=langViews%><br></b><select name="view" multiple size="10">
<%		for each item in extDba.Views.Items%>
					<option value="<%=item.Name%>"><%=item.Name%></option>
<%		next%>
				</select></td>
				
				<td><b><%=langProcedures%><br></b><select name="procedure" multiple size="10">
<%		for each item in extDba.Procedures.Items%>
					<option value="<%=item.Name%>"><%=item.Name%></option>
<%		next%>
				</select></td>
				
				<td valign="top">
					<b><%=langOptions%></b><br>
					<input type="checkbox" name="relations" value="1">&nbsp;<%=langIncludeRelations%><br>
					<input type="checkbox" name="data" value="1">&nbsp;<%=langIncludeData%><br>
					<input type="checkbox" name="transaction" value="1">&nbsp;<%=langUseTransaction%><br>
					<input type="checkbox" name="ignore_errors" value="1" checked>&nbsp;<%=langIgnoreErrors%>
				</td>
			</tr>
			<tr>
				<td colspan="4" align="center">
					<input type="button" value="<< <%=langBack%>" class="button" onclick="javascript:window.location.href='import_db.asp?e_path=<%=Server.URLEncode(Request.Form("e_path").Item)%>';">&nbsp;
					<input type="submit" value="<%=langImport%>" class="button">
				</td>
			</tr>
		</table>
		</form>
<%		
		set extDba = Nothing
		call DBA_EndNewTable
	End Sub
%>

<%
	Sub DoImport
		dim extPath, extPwd, extDba, tbl, key, strErrors, fld, tblNew
		Dim strTables, bWithData, strAutoNumber, recOld, recNew, i, UseTransaction, IgnoreErrors
		extPath = Request.Form("e_path").Item
		extPwd = Request.Form("e_pwd").Item
		strErrors = ""
		strTables = ""
		if Request.Form("data").Item = "1" then bWithData = True else bWithData = False
		If Request.Form("transaction").Item = "1" Then UseTransaction = True Else UseTransaction = False
		If Request.Form("ignore_errors").Item = "1" Then IgnoreErrors = True Else IgnoreErrors = False
		set extDba = new DBAdmin
		call extDba.Connect(Request.Form("e_path").Item, Request.Form("e_pwd").Item)
		
		If UseTransaction Then call dba.JetConnection.BeginTrans()
		
		'start importing tables
		For Each key In Request.Form("table")
			set tbl = extDba.Tables.Item(key)
			call dba.CreateTable(tbl.Name)
			if dba.HasError then 
				strErrors = strErrors & dba.LastError & "<br>"
				If not IgnoreErrors Then Exit For
			else
				set tblNew = dba.Tables.Item(tbl.Name)
				if TypeName(tblNew) <> "Nothing" Then
					strAutoNumber = ""
					For each fld in tbl.Fields.Items
						If fld.IsAutoNumber then strAutoNumber = fld.Name
						call tblNew.CreateField(fld, False)
						If dba.HasError Then strErrors = strErrors & dba.LastError
					Next
					If not IgnoreErrors and Len(strErrors) > 0 Then Exit For

					For Each fld in tbl.Indexes.Items
						If not fld.IsForeignKey then call tblNew.CreateIndex(fld.Name, fld.TargetField, fld.IsUnique, fld.IsPrimary)
						If dba.HasError Then strErrors = strErrors & dba.LastError
					Next
					If not IgnoreErrors and Len(strErrors) > 0 Then Exit For

					'import data now
					if bWithData Then
						set recOld = tbl.GetRawData(10, "", True)
						set recNew = tblNew.GetRawData(10, "", False)
						Do While not recOld.EOF
							call recNew.AddNew
							For i=0 To recOld.Fields.Count - 1
								'some more checks need to be done in case of Null values
								'I HATE Null values! :)
								if IsNull(recOld(i).Value) Then
									set fld = tbl.Fields.Item(recOld(i).Name)
									if fld.IsNullable then
										recNew(recOld(i).Name).Value = recOld(i).Value
									elseif fld.HasDefault then
										'just leave it as is
									elseif fld.FieldType = 203 or fld.FieldType = 130 and fld.AllowZeroLength Then
										recNew(recOld(i).Name).Value = ""
									else
										recNew(recOld(i).Name).Value = 0
									End If
								else
									recNew(recOld(i).Name).Value = recOld(i).Value
								end if
							Next
							call recNew.Update
							If dba.HasError Then strErrors = strErrors & dba.LastError
							If not IgnoreErrors and Len(strErrors) > 0 Then Exit Do
							call recOld.MoveNext
						Loop
						recNew.Close
						recOld.Close
						set recNew = Nothing
						set recOld = Nothing
					End If
					
					strTables = strTables & tbl.Name & "!"
					set tblNew = Nothing
				End If
			end if
			set tbl = Nothing
			If not IgnoreErrors and Len(strErrors) > 0 Then Exit For
		Next
		
		'import views now
		If IgnoreErrors or Len(strErrors) = 0 Then
			For Each key In Request.Form("view")
				set tbl = extDba.Views.Item(key)
				If TypeName(tbl) <> "Nothing" then
					call dba.CreateView(tbl.Name, tbl.Body)
					if dba.HasError then strErrors = strErrors & dba.LastError & "<br>"
				End If
				set tbl = Nothing	
			Next
		End If
		
		'import stored procedures now
		If IgnoreErrors or Len(strErrors) = 0 Then
			For Each key In Request.Form("procedure")
				set tbl = extDba.Procedures.Item(key)
				If TypeName(tbl) <> "Nothing" then
					call dba.CreateProcedure(tbl.Name, tbl.Body)
					if dba.HasError then strErrors = strErrors & dba.LastError & "<br>"
				End If
				set tbl = Nothing	
			Next
		End If
		
		'Copy Relations now
		If IgnoreErrors or Len(strErrors) = 0 Then
			If Request.Form("relations").Item = "1" Then
				For Each tbl In extDba.Relations.Items
					If InStr(1, strTables, tbl.PrimaryTable) > 0 and InStr(1, strTables, tbl.ForeignTable) > 0 Then
						call dba.CreateRelation( _
							tbl.Name, _
							tbl.PrimaryTable, _
							tbl.PrimaryField, _
							tbl.ForeignTable, _
							tbl.ForeignField, _
							tbl.OnUpdate, _
							tbl.OnDelete _
						)
						if dba.HasError then strErrors = strErrors & dba.LastError & "<br>"
					End If
				Next
			End If
		End If
		
		set extDba = Nothing
		
		If UseTransaction and (IgnoreErrors or Len(strErrors) = 0) Then
			call dba.JetConnection.CommitTrans()
		ElseIf UseTransaction Then
			call dba.JetConnection.RollbackTrans()
		End If
		
		'finished importing now.. Let's see errors
		call DBA_BeginNewTable(langImportDatabase, "", "90%", "")
		If Len(strErrors) > 0 then
			call DBA_WriteError(strErrors)
		else
			call DBA_WriteSuccess(langImportSuccess)
		end if
		call DBA_EndNewTable
	End Sub
%>