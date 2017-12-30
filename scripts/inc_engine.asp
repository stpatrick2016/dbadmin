<%


'//////////////////////////////////////////////////////////////////////////////////////////////////
'// Stp Database Administrator Engine
'// Engine version: 1.1
'// Copyright © 2002-2003 by Philip Patrick. All rights reserved
'//
'// Author:		Philip Patrick
'// E-mail:		stpatrick@mail.com
'// Web-site:	http://www.stpworks.com
'// Description:
'//		Set of classes and functions for managing Access database on the Web

Const DBAE_JET_PROVIDER		= "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Const DBAE_DEBUG			= False


'/////////////////////////////////////////////////////////
'// Global functions

'#Makes syntax coloring for given SQL statement
Function HighlightSQL(sSQL)
	Const KeyWords =	"CREATE|TABLE|COUNTER|NOT NULL|DEFAULT|INDEX|ON|PRIMARY|WITH|LONG|TEXT|DATETIME|BIT|MONEY|BINARY|TINYINT|DECIMAL|FLOAT|INTEGER|REAL|UNIQUEIDENTIFIER|MEMO|UNIQUE|INSERT|INTO|SELECT|FROM|WHERE|UPDATE|DELETE|VALUES|PARAMETERS|ORDER BY|OR|AND|IN|SUM|AS|TOP|SET|LEFT|RIGHT|INNER|JOIN|ASC|DESC|GROUP BY|HAVING|CONSTRAINT|ADD|COLUMN|CASCADE|DROP|TOP|DISTINCT|DISTINCTROW|KEY|MIN|MAX|COUNT|AVG|PROCEDURE|VIEW|STDEV|STDEVP|UNION|ALTER|REFERENCES|FOREIGN|NO ACTION"
	
	dim RegEx, s
	set RegEx = new RegExp
	RegEx.Global = True
	RegEx.IgnoreCase = true
	
	sSQL = Replace(sSQL, vbCrLf, "<br>")
	
	'Replace code
	RegEx.Pattern = "(\b" & Replace(KeyWords, "|", "\b|\b") & "\b)"
	sSQL = RegEx.Replace(sSQL, "<font color=""blue"">$1</font>")
	
	'replace numbers
	RegEx.Pattern = "([\s\(<>=\-\+])([0-9]+)([\s,;\)<>=\-\+])"
	sSQL = RegEx.Replace(sSQL, "$1<font color=""green"">$2</font>$3")
	
	set RegEx = nothing
	HighlightSQL = sSQL
End Function

'/////////////////////////////////////////////////////////
'// Classes
Class DBAdmin

	'constructor
	Private Sub Class_Initialize
		Set Tables_		= Server.CreateObject("Scripting.Dictionary")
		Set Views_		= Server.CreateObject("Scripting.Dictionary")
		Set Relations_	= Server.CreateObject("Scripting.Dictionary")
		Set Procedures_	= Server.CreateObject("Scripting.Dictionary")
		
		EngineVersion_	= "1.2"
		
		call Reset
	End Sub

	'destructor
	Private Sub Class_Terminate
		call Reset

		Set Tables_		= Nothing
		Set Views_		= Nothing
		Set Relations_	= Nothing
		Set Procedures_	= Nothing
	End Sub


	'######################################################## 
	'#Returns the version of Engine (not the whole product)
	Public Property Get EngineVersion
		EngineVersion = EngineVersion_
	End Property  

	'######################################################## 
	'#Path to Access database
	Public Property Let DatabasePath(v)
		call Reset
		DatabasePath_ = CStr(v)
	End Property    
	
	Public Property Get DatabasePath
		DatabasePath = DatabasePath_
	End Property  

	'######################################################## 
	'#Active ADO Connection object
	Public Property Get JetConnection
		Set JetConnection = JetConnection_
	End Property  

	'######################################################## 
	'#Last error occured in operation
	Public Property Let LastError(v)
		LastError_ = CStr(v)
	End Property    
	
	Public Property Get LastError
		LastError = LastError_
	End Property  

	'######################################################## 
	'#Returns a size of database file in bytes
	Public Property Get Size
		Size = 0
		
		dim fso, f
		if not DBAE_DEBUG then On Error Resume Next
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		set f = fso.GetFile(DatabasePath_)
		
		If not IsError then
			Size = f.Size
		end if

		set f = Nothing
		set fso = Nothing
	End Property  

	'######################################################## 
	'#Returns how much space can be reclaimed after compacting the database
	Public Property Get ReclaimedSpace
		ReclaimedSpace = 0
		
		If not DBAE_DEBUG then On Error Resume Next
		If IsOpen then
			ReclaimedSpace = CLng(JetConnection_.Properties("Jet OLEDB:Compact Reclaimed Space Amount").Value)
		end if
	End Property  

	'######################################################## 
	'# Returns locale identifier of the database
	Public Property Get LocaleIdentifier
		If not DBAE_DEBUG then On Error Resume Next
		
		If IsOpen then LocaleIdentifier = JetConnection_.Properties("Locale Identifier").Value
	End Property

	'######################################################## 
	'#Dictionary object contains all tables in database
	Public Property Get Tables
		if Tables_.Exists(".uninitialized") then
			'first time. Let's get tables names
			dim tbl, xTable, xCat
			Tables_.RemoveAll
			if not DBAE_DEBUG then On Error Resume Next
			set xCat = Server.CreateObject("ADOX.Catalog")
			if xCat Is Nothing or IsEmpty(xCat) Then
				'ADOX is not available, so we'll get tables list using schemas
				set xCat = JetConnection_.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
				Do While Not xCat.EOF
					set tbl = new DBATable
					With tbl
						.Name = xCat("TABLE_NAME").Value
						.DateCreated = xCat("DATE_CREATED").Value
						.DateModified = xCat("DATE_MODIFIED").Value
						.Description = xCat("DESCRIPTION").Value
						Set .Parent = Me
					End With
					Set Tables_.Item(tbl.Name) = tbl
					xCat.MoveNext
				Loop
				call xCat.Close()
			Else
				set xCat.ActiveConnection = JetConnection_
				if IsError then Exit Property
				for each xTable in xCat.Tables
					if xTable.Type = "TABLE" then 
						set tbl = new DBATable
						with tbl
							.Name = xTable.Name
							.DateCreated = xTable.DateCreated
							.DateModified = xTable.DateModified
							.Description = ""
							Set .Parent = Me
						end with
						Set Tables_.Item(tbl.Name) = tbl
					end if
				next
				
			End If
			set xCat = nothing
		end if
		
		Set Tables = Tables_
	End Property  

	'######################################################## 
	'#Dictionary object contains all procedures in database
	Public Property Get Procedures
		if Procedures_.Exists(".uninitialized") then
			dim p, xProc, xCat
			Procedures_.RemoveAll
			
			if not DBAE_DEBUG then On Error Resume Next
			set xCat = Server.CreateObject("ADOX.Catalog")
			if IsEmpty(xCat) or xCat is Nothing Then
				set xCat = JetConnection_.OpenSchema(adSchemaProcedures)
				Do While Not xCat.EOF
					set p = new DBAProcedure
					With p
						.Name = xCat("PROCEDURE_NAME").Value
						.Body = xCat("PROCEDURE_DEFINITION").Value
						.DateCreated = xCat("DATE_CREATED").Value
						.DateModified = xCat("DATE_MODIFIED").Value
						.Description = xCat("DESCRIPTION").Value
						Set .Parent = Me
					End With
					Set Procedures_.Item(p.Name) = p
					xCat.MoveNext
				Loop
				xCat.Close
			Else
				set xCat.ActiveConnection = JetConnection_
				If IsError Then Exit Property
				for each xProc in xCat.Procedures
					set p = new DBAProcedure
					with p
						.Name = xProc.Name
						.Body = xProc.Command.CommandText
						.DateCreated = xProc.DateCreated
						.DateModified = xProc.DateModified
						.Description = ""
						Set .Parent = Me
					end with
					Set Procedures_.Item(p.Name) = p
				next
			End If
			
			set xCat = nothing
		end if
	
		Set Procedures = Procedures_
	End Property  

	'######################################################## 
	'# Dictionary object contains all views in database
	Public Property Get Views
		if Views_.Exists(".uninitialized") then
			dim v, xCat, xView
			Views_.RemoveAll
			
			if not DBAE_DEBUG then On Error Resume Next
			set xCat = Server.CreateObject("ADOX.Catalog")
			if IsEmpty(xCat) or xCat Is Nothing Then
				set xCat = JetConnection_.OpenSchema(adSchemaViews)
				Do While Not xCat.EOF
					set v = new DBAView
					With v
						.Name = xCat("TABLE_NAME").Value
						.Body = xCat("VIEW_DEFINITION").Value
						.DateCreated = xCat("DATE_CREATED").Value
						.DateModified = xCat("DATE_MODIFIED").Value
						.Description = xCat("DESCRIPTION").Value
						Set .Parent = Me
					End With
					Set Views_.Item(v.Name) = v
					xCat.MoveNext
				Loop
				xCat.Close
			Else
				set xCat.ActiveConnection = JetConnection_
				If IsError Then Exit Property
				for each xView in xCat.Views
					set v = new DBAView
					with v
						.Name = xView.Name
						.Body = xView.Command.CommandText
						.DateCreated = xView.DateCreated
						.DateModified = xView.DateModified
						.Description = ""
						Set .Parent = Me
					end with
					Set Views_.Item(v.Name) = v
				next
			End If
			
			set xCat = Nothing
		end if
		
		Set Views = Views_
	End Property  

	'######################################################## 
	'# Dictionary Object contains all relationships in database
	Public Property Get Relations
		if Relations_.Exists(".uninitialized") then
			dim rec, rel
			Relations_.RemoveAll
			
			if not DBAE_DEBUG then On Error Resume Next
			set rec = JetConnection_.OpenSchema(adSchemaForeignKeys)
			If IsError Then Exit Property
			do while not rec.EOF
				set rel = new DBARelation
				with rel
					.Name = rec("FK_NAME").Value
					.PrimaryTable = rec("PK_TABLE_NAME").Value
					.PrimaryField = rec("PK_COLUMN_NAME").Value
					.PrimaryIndex = rec("PK_NAME").Value
					.ForeignTable = rec("FK_TABLE_NAME").Value
					.ForeignField = rec("FK_COLUMN_NAME").Value
					.OnUpdate = rec("UPDATE_RULE").Value
					.OnDelete = rec("DELETE_RULE").Value
					Set .Parent = Me
				end with
				Set Relations_.Item(rel.Name) = rel
				
				rec.MoveNext
			loop
			rec.Close
			set rec = nothing
		end if

		Set Relations = Relations_
	End Property  

	'######################################################## 
	'# Returns True if the database is Access 97 database
	Public Property Get IsAccess97
		if not DBAE_DEBUG then On Error Resume Next
		IsAccess97 = False
		if IsOpen then
			if CInt(JetConnection_.Properties("Jet OLEDB:Engine Type")) = 5 then IsAccess97 = False else IsAccess97 = True
		end if
	End Property

	'######################################################## 
	'# Opens a database connection, closing the existing one is present
	Public Function Connect(MDBPath, Password)
		dim strCon

		Connect = True
		call Reset
		
		'check if DSN was passed and retrieve file name
		if InStr(1, MDBPath, "DSN=", vbTextCompare) = 1 then MDBPath = GetFilenameFromDSN(Mid(MDBPath, 5), Password)
		
		DatabasePath_ = CStr(MDBPath)
		DatabasePassword_ = CStr(Password)
		
		strCon = DBAE_JET_PROVIDER & DatabasePath_
		if Len(DatabasePassword_) > 0 then strCon = strCon & ";Jet OLEDB:Database password=" & DatabasePassword_
		Set JetConnection_ = Server.CreateObject("ADODB.Connection")
		
		if not DBAE_DEBUG then On Error Resume Next
		JetConnection_.Open strCon
		
		if IsError then
			dim lastErr : lastErr = LastError
			call Reset
			LastError = lastErr
			Connect = False
		end if
	End Function

	'######################################################## 
	'# Creates a new blank database, and if successful, opens current connection with it
	Public Function CreateDatabase(Path)
		dim catalog
		
		if not DBAE_DEBUG then On Error Resume Next
		set catalog = Server.CreateObject("ADOX.Catalog")
		if IsEmpty(catalog) or catalog Is Nothing Then
			LastError = "ADOX is not available. Database couldn't be created"
		Else
			if Right(Path, 4) <> ".mdb" then Path = Path & ".mdb"
			
			call catalog.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Path)
			
			set catalog = nothing
			
			if not IsError then call Connect(Path, "")
		End If
		CreateDatabase = not HasError
	End Function

	'######################################################## 
	'# Creates a new table in existing database
	Public Function CreateTable(Name)
		If not IsOpen then
			LastError = "Object is not initialized"
			CreateTable = False
			Exit Function
		end if
		
		dim objTbl
		if not DBAE_DEBUG then On Error Resume Next
		JetConnection_.Execute "CREATE TABLE [" & Name & "]", adExecuteNoRecords
		
		CreateTable = not IsError
		if Len(LastError) = 0 Then
			'remove all tables and reload them
			Tables_.Item(".uninitialized") = null
		end if
	End Function

	'######################################################## 
	'# Deletes an existing table in database
	Public Function DeleteTable(Name)
		If not IsOpen then
			LastError = "Object is not initialized"
			DeleteTable = False
			Exit Function
		end if
		
		if not DBAE_DEBUG then On Error Resume Next
		JetConnection_.Execute "DROP TABLE [" & Name & "]", adExecuteNoRecords
		
		DeleteTable = not IsError
		if Len(LastError) = 0 Then
			'delete table from tables list
			if Tables_.Exists(Name) then Tables_.Remove Name
		end if
	End Function

	'######################################################## 
	'# Creates a new stored procedure
	Public Function CreateProcedure(Name, Body)
		If not IsOpen then
			LastError = "Object is not initialized"
			CreateProcedure = False
			Exit Function
		end if
		
		dim xCat, cmd
		if not DBAE_DEBUG then On Error Resume Next
		set xCat = Server.CreateObject("ADOX.Catalog")
		If IsEmpty(xCat) or xCat Is Nothing Then
			Err.Clear
			cmd = "CREATE PROCEDURE [" & Name & "] AS " & Body
			call JetConnection_.Execute(cmd, adExecuteNoRecords)
		Else
			set xCat.ActiveConnection = JetConnection_
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.CommandText = Body
			call xCat.Procedures.Append(Name, cmd)
			
			set cmd = Nothing
			set xCat = Nothing
		End If
		CreateProcedure = not IsError
		if not HasError then
			Procedures_.Item(".uninitialized") = null
		end if
	End Function

	'######################################################## 
	'# Deletes an existing stored procedure
	Public Function DeleteProcedure(Name)
		If not IsOpen then
			LastError = "Object is not initialized"
			DeleteProcedure = False
			Exit Function
		end if
		
		if not DBAE_DEBUG then On Error Resume Next
		JetConnection_.Execute "DROP PROCEDURE [" & Name & "]", adExecuteNoRecords
		
		DeleteProcedure = not IsError
		if not HasError and Procedures_.Exists(Name) then Procedures_.Remove Name
	End Function

	'######################################################## 
	'# Creates a new view
	Public Function CreateView(Name, Body)
		If not IsOpen then
			LastError = "Object is not initialized"
			CreateView = False
			Exit Function
		end if
		
		dim xCat, cmd
		if not DBAE_DEBUG then On Error Resume Next
		set xCat = Server.CreateObject("ADOX.Catalog")
		If IsEmpty(xCat) or xCat Is Nothing Then
			Err.Clear
			cmd = "CREATE PROCEDURE [" & Name & "] AS " & Body
			call JetConnection_.Execute(cmd, adExecuteNoRecords)
		Else
			set xCat.ActiveConnection = JetConnection_
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.CommandText = Body
			call xCat.Views.Append(Name, cmd)
			
			set cmd = Nothing
			set xCat = Nothing
		End If
		CreateView = not IsError
		if not HasError then
			Views_.Item(".uninitialized") = null
		end if
	End Function

	'######################################################## 
	'# Deletes an existing view
	Public Function DeleteView(Name)
		If not IsOpen then
			LastError = "Object is not initialized"
			DeleteView = False
			Exit Function
		end if
		
		if not DBAE_DEBUG then On Error Resume Next
		JetConnection_.Execute "DROP VIEW [" & Name & "]", adExecuteNoRecords
		
		DeleteView = not IsError
		if Len(LastError) = 0 then 
			if Views_.Exists(Name) then Views_.Remove Name
		end if
	End Function

	'######################################################## 
	'# Creates a new relationship
	Public Function CreateRelation(Name, PKTable, PKField, FKTable, FKField, OnUpdate, OnDelete)
		If not IsOpen then
			LastError = "Object is not initialized"
			CreateRelation = False
			Exit Function
		end if
		
		dim sSQL
		sSQL =	"ALTER TABLE [" & FKTable & "] ADD CONSTRAINT [" &_
				Name & "] FOREIGN KEY ([" & FKField &_
				"]) REFERENCES [" & PKTable & "]([" &_
				PKField & "])"
		if Len(OnUpdate) > 0 then sSQL = sSQL & " ON UPDATE " & OnUpdate
		if Len(OnDelete) > 0 then sSQL = sSQL & " ON DELETE " & OnDelete
		
		if not DBAE_DEBUG then On Error Resume Next
		JetConnection_.Execute sSQL, adExecuteNoRecords
		
		CreateRelation = not IsError
		if Len(LastError) = 0 then 
			Relations_.Item(".uninitialized") = null
		end if
	End Function
	
	'######################################################## 
	'# Deletes an existing relationship
	Public Function DeleteRelation(Name, FKTable)
		If not IsOpen then
			LastError = "Object is not initialized"
			DeleteRelation = False
			Exit Function
		end if
		
		dim sSQL
		sSQL =	"ALTER TABLE [" & FKTable & "] DROP CONSTRAINT [" &_
				Name & "]"
		
		if not DBAE_DEBUG then On Error Resume Next
		JetConnection_.Execute sSQL, adExecuteNoRecords
		
		DeleteRelation = not IsError
		if Len(LastError) = 0 then 
			if Relations_.Exists(Name) then Relations_.Remove Name
		end if
	End Function

	'######################################################## 
	'# Compacts and repaires a database. Converts Access 97 databases to Access 2000
	'# If new password not null, then changes/sets a new password to database
	Public Function CompactDatabase(DoUpgrade, NewPassword, NewLocaleID)
		If not IsOpen then
			LastError = "Object is not initialized"
			CompactDatabase = False
			Exit Function
		end if
		
		dim strTempFile, fso, jro, ver, strCon, strTo, LCID
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		
		strTempFile = DatabasePath_
		strTempFile = Left(strTempFile, InStrRev(strTempFile, "\")) & fso.GetTempName
		set jro = Server.CreateObject("JRO.JetEngine")
		if not DoUpgrade and IsAccess97 then ver = "4" else ver = "5"
		
		'close the database first
		if Len(NewLocaleID) > 0 Then LCID = NewLocaleID Else LCID = JetConnection_.Properties("Locale Identifier").Value
		JetConnection_.Close

		strCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & DatabasePath_
		if Len(DatabasePassword_) > 0 then strCon = strCon & ";Jet OLEDB:Database password=" & DatabasePassword_
		strTo = "Provider=Microsoft.Jet.OLEDB.4.0; Locale Identifier=" & LCID & "; Data Source=" & strTempFile & "; Jet OLEDB:Engine Type=" & ver
		if Len(DatabasePassword_) > 0 and IsNull(NewPassword) then 
			strTo = strTo & ";Jet OLEDB:Database password=" & DatabasePassword_
		elseif not IsNull(NewPassword) and Len(NewPassword) > 0 then
			strTo = strTo & ";Jet OLEDB:Database password=" & NewPassword
		end if
		
		if not DBAE_DEBUG then On Error Resume Next
		jro.CompactDatabase strCon, strTo

		CompactDatabase = False
		if IsError then
			fso.DeleteFile strTempFile
		else
			fso.DeleteFile DatabasePath_
			fso.MoveFile strTempFile, DatabasePath_
			if IsError then
				fso.DeleteFile strTempFile
			else
				CompactDatabase = True
				if not IsNull(NewPassword) then DatabasePassword_ = NewPassword
			end if
		end if
		set jro = nothing
		set fso = nothing
		
		'reopen the database
		strCon = DBAE_JET_PROVIDER & DatabasePath_
		if Len(DatabasePassword_) > 0 then strCon = strCon & ";Jet OLEDB:Database password=" & DatabasePassword_
		JetConnection_.Open strCon
	End Function
	
	'######################################################## 
	'# Creates a backup copy of opened database
	Public Function BackupDatabase()
		If not IsOpen then
			LastError = "Object is not initialized"
			BackupDatabase = False
			Exit Function
		end if
	
		dim fso, sFileName
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		sFileName = DatabasePath_
		sFileName = Left(sFileName, InStrRev(sFileName, ".")) & "bak"
		
		'close the database first
		JetConnection_.Close
		
		if not DBAE_DEBUG then On Error Resume Next
		fso.CopyFile DatabasePath_, sFileName, True
		
		BackupDatabase = not IsError
		set fso = nothing
		
		'reopen the database
		dim strCon
		strCon = DBAE_JET_PROVIDER & DatabasePath_
		if Len(DatabasePassword_) > 0 then strCon = strCon & ";Jet OLEDB:Database password=" & DatabasePassword_
		JetConnection_.Open strCon
	End Function

	'######################################################## 
	'# Restores a database from previously created backup copy
	Public Function RestoreDatabase()
		If not IsOpen then
			LastError = "Object is not initialized"
			RestoreDatabase = False
			Exit Function
		end if

		dim fso, sFileName
		
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		sFileName = DatabasePath_
		sFileName = Left(sFileName, InStrRev(sFileName, ".")) & "bak"
		
		'close the database first
		JetConnection_.Close
		
		if not DBAE_DEBUG then On Error Resume Next
		fso.CopyFile sFileName, DatabasePath_, True
		
		RestoreDatabase = not IsError
		set fso = nothing
		
		'reopen the database
		dim strCon
		strCon = DBAE_JET_PROVIDER & DatabasePath_
		if Len(DatabasePassword_) > 0 then strCon = strCon & ";Jet OLEDB:Database password=" & DatabasePassword_
		JetConnection_.Open strCon
	End Function

	'######################################################## 
	'# Returns True if the object is initialized
	Public Function IsOpen()
		if IsObject(JetConnection_) and Len(DatabasePath_) > 0 Then IsOpen = True Else IsOpen = False
	End Function

	'######################################################## 
	'# Returns True if any error occured
	Public Function HasError()
		if Len(LastError) > 0 Then HasError = True Else HasError = False
	End Function
	
	'######################################################## 
	'# Resets the object to uninitialized state
	Public Sub Reset()
		Tables_.RemoveAll
		Relations_.RemoveAll
		Views_.RemoveAll
		Procedures_.RemoveAll
		Tables_.Add ".uninitialized", null
		Relations_.Add ".uninitialized", null
		Views_.Add ".uninitialized", null
		Procedures_.Add ".uninitialized", null
		
		DatabasePath_		= ""
		LastError_			= ""
		DatabasePassword_	= ""
		
		if IsObject(JetConnection_) Then
			On Error Resume Next
			JetConnection_.Close
			Set JetConnection_ = Nothing
		end if
	End Sub

	'######################################################## 
	'# Checks and update last error
	Function IsError
		If Err then
			LastError = Err.Description & " (" & Err.number & ")"
			IsError = True
			Err.Clear
		else
			LastError = ""
			IsError = False
		end if
	End Function

	'######################################################## 
	'# Executes a script, which can consist of several SQL statements, separated
	'# with ";".
	'# Transaction (Boolean) means run the script as one transaction
	'# IgnoreErrors (Boolean) - finish the script regardless any errors that may occur
	Public Function RunScript(Script, Transaction, IgnoreErrors, ByRef ArrayAffected)
		dim arrSQL, q, rec, intAffected, con, strCon, i, re
		
		if not DBAE_DEBUG then On Error Resume Next
		LastError_ = ""
		if Not IsNull(ArrayAffected) then Redim ArrayAffected(-1)
		Transaction = CBool(Transaction)
		IgnoreErrors = CBool(IgnoreErrors)
		
		'create a new connection object - for adUseClient
		set con = Server.CreateObject("ADODB.Connection")
		con.CursorLocation = adUseClient
		strCon = DBAE_JET_PROVIDER & DatabasePath_
		if Len(DatabasePassword_) > 0 then strCon = strCon & ";Jet OLEDB:Database password=" & DatabasePassword_
		con.Open strCon
		if IsError then Exit Function
		
		if Transaction then call con.BeginTrans

		arrSQL = Split(Script, ";")
		set re = new RegExp
		re.Pattern = "create\s+procedure(.|\n)+parameters(\w|\s)*$"
		re.IgnoreCase = True
		for i=0 to ubound(arrSQL)
			q = arrSQL(i)
			'since Trim doesn't remove vbCrLf from its own reason, then I will delete it
			q = Replace(q, vbCrLf, " ")
			q = Trim(q)
			if re.Test(q) and i < ubound(arrSQL) then
				arrSQL(i+1) = q & "; " & arrSQL(i+1)
				q = ""
			end if
			if Len(q) > 0 then
				set rec = con.Execute(q, intAffected)
				if not IsNull(ArrayAffected) then 
					Redim Preserve ArrayAffected(ubound(ArrayAffected) + 1)
					ArrayAffected(ubound(ArrayAffected)) = CInt(intAffected)
				end if
				if Err then
					LastError_ = LastError_ & Err.Description & vbCrLf
					if not IgnoreErrors then Exit For
					Err.Clear
				end if
			end if
		next
		set re = Nothing
		
		if Transaction and HasError and not IgnoreErrors then
			call con.RollbackTrans
		elseif Transaction then
			call con.CommitTrans
		end if
		
		If not IsObject(rec) then 
			set rec = Server.CreateObject("ADODB.Recordset")
		end if
		
		'detach from connection object
		rec.ActiveConnection = Nothing
		
		con.Close
		set con = Nothing

		set RunScript = rec
	End Function
	
	'---------------------------
	'protected and private


	Private Tables_
	Private Relations_
	Private Views_
	Private Procedures_
	Private DatabasePath_
	Private DatabasePassword_
	Private JetConnection_
	Private LastError_
	Private EngineVersion_
	
	Private Function GetFilenameFromDSN(dsnName, pwd)
		dim dsn, ret, i
		ret = ""
		set dsn = Server.CreateObject("ADODB.Connection")

		if not DBAE_DEBUG then On Error Resume Next
		call dsn.Open("DSN=" & dsnName, "Admin", pwd)
		if not IsError then
			ret = dsn.Properties("Current Catalog").Value
			if Len(ret) > 0 then
				if Right(ret, 4) <> ".mdb" then ret = ret & ".mdb"
			else
				ret = dsn.Properties("Extended Properties").Value
				i = InStr(1, ret, "DBQ=", vbTextCompare)
				if i > 0 then
					ret = Left(ret, i+4)
					i = InStr(1, ret, ";")
					ret = Left(ret, i-1)
				else
					ret = ""
				end if
			end if
			dsn.Close
		end if
		set dsn = Nothing
		GetFilenameFromDSN = ret
	End Function

End Class
' END CLASS DEFINITION DBAdmin


'///////////////////////////////////////////////////////////////
'// Holds any information of the table. When this class first created it contains only Name, but when any other property is accessed, it calls Reload to load all other information from database
'//
Class DBATable

	'constructor
	Private Sub Class_Initialize
		Set Fields_		= Server.CreateObject("Scripting.Dictionary")
		Set Indexes_	= Server.CreateObject("Scripting.Dictionary")
		Fields_.Add ".uninitialized", null
		Indexes_.Add ".uninitialized", null

		Name_			= ""
		Description_	= ""
		DateCreated_	= null
		DateModified_	= null
		Set Parent_		= Nothing
	End Sub
	
	'destructor
	Private Sub Class_Terminate
		Fields_.RemoveAll
		Indexes_.RemoveAll
		Set Fields_		= Nothing
		Set Indexes_	= Nothing
	End Sub



	'######################################################## 
	'# Parent object - DBAdmin
	Public Property Set Parent(v)
		if IsObject(v) then Set Parent_ = v
	End Property
	
	Public Property Get Parent
		if IsObject(Parent_) then Set Parent = Parent_ else Set Parent = Nothing
	End Property  

	'######################################################## 
	'# Name of the table
	Public Property Let Name(v)
		if Len(Name_) = 0 then Name_ = CStr(v)
	End Property    
	
	Public Property Get Name
		Name = Name_
	End Property  

	'######################################################## 
	'# Fields collection
	Public Property Get Fields
		if not IsInitialized then Exit Property
		
		if Fields_.Exists(".uninitialized") then
			dim rec, f, xCat, bNoADOX
			Fields_.RemoveAll
			
			if not DBAE_DEBUG then On Error Resume Next
			set rec = Parent_.JetConnection.OpenSchema(adSchemaColumns, Array(empty,empty, Name_))
			set xCat = Server.CreateObject("ADOX.Catalog")
			if (IsEmpty(xCat) or xCat Is Nothing) Then
				Err.Clear
				set xCat = Parent_.JetConnection.Execute(Name_)
				bNoADOX = True
			else
				set xCat.ActiveConnection = Parent_.JetConnection
				bNoADOX = False
			End If
			If Parent_.IsError then exit Property
			do while not rec.EOF
				set f = new DBAField
				with f
					.Name = rec("COLUMN_NAME").Value
					if bNoADOX Then .FieldType = rec("DATA_TYPE").Value else .FieldType = xCat.Tables(Name_).Columns(.Name).Type
					.MaxLength = rec("CHARACTER_MAXIMUM_LENGTH").Value
					.DefaultValue = rec("COLUMN_DEFAULT").Value
					.IsNullable = rec("IS_NULLABLE").Value
					.Ordinal = rec("ORDINAL_POSITION").Value
					.Description = rec("DESCRIPTION").Value
					if bNoADOX Then
						.IsAutonumber = xCat(.Name).Properties("ISAUTOINCREMENT").Value
						.Compressed = False
						.AllowZeroLength = False
					Else
						.IsAutonumber = xCat.Tables(Name_).Columns(.Name).Properties("Autoincrement").Value
						.Compressed = xCat.Tables(Name_).Columns(.Name).Properties("Jet OLEDB:Compressed UNICODE Strings").Value
						.AllowZeroLength = xCat.Tables(Name_).Columns(.Name).Properties("Jet OLEDB:Allow Zero Length").Value
					End If
					Set .Parent = Me
				end with
				Set Fields_.Item(f.Name) = f

				rec.MoveNext
			loop
			rec.Close
			if bNoADOX Then xCat.Close
			set rec = nothing
			set xCat = Nothing
		end if
		
		Set Fields = Fields_
	End Property  

	'######################################################## 
	'# Indexes collection
	Public Property Get Indexes
		if not IsInitialized then Exit Property
		
		if Indexes_.Exists(".uninitialized") then
			dim rec, indx
			Indexes_.RemoveAll
			
			if not DBAE_DEBUG then On Error Resume Next
			set rec = Parent_.JetConnection.OpenSchema(adSchemaIndexes,Array(empty,empty,empty,empty, Name_))
			If Parent_.IsError then Exit Property
			do while not rec.EOF
				set indx = new DBAIndex
				with indx
					.Name = rec("INDEX_NAME").Value
					.TargetField = rec("COLUMN_NAME").Value
					.IsUnique = rec("UNIQUE").Value
					.IsPrimary = rec("PRIMARY_KEY").Value
					Set .Parent = Me
				end with
				Set Indexes_.Item(indx.Name & "." & indx.TargetField) = indx
				
				rec.MoveNext
			loop
			rec.Close
			set rec = nothing
		end if
		
		Set Indexes = Indexes_
	End Property  

	'######################################################## 
	'# Contains SQL statment for creating this table, including indexes, but not including relationships
	Public Property Get SQL
		dim strSQL, strTemp, item
		strSQL = "CREATE TABLE [" & Name_ & "]"
		
		'get fields definitions
		strTemp = ""
		for each item in Fields.Items
			strTemp = strTemp & item.SQL & ", "
		next
		if Len(strTemp) > 0 then 
			strTemp = Left(strTemp, Len(strTemp) - 2)
			strSQL = strSQL & "(" & strTemp & ")"
		end if
		strSQL = strSQL & ";" & vbCrLf & vbCrLf
		
		'get all indexes
		strTemp = ""
		for each item in Indexes.Items
			if InStr(1, strTemp, item.Name, vbTextCompare) <= 0 and not item.IsForeignKey then
				strSQL = strSQL & item.SQL & ";" & vbCrLf
				strTemp = strTemp & item.Name & "."
			end if
		next
		
		SQL = strSQL
	End Property  

	'######################################################## 
	'# Read-only value of description of the table
	Public Property Get Description
		Description = Description_
	End Property  

	Public Property Let Description(v)
		if Len(Description_) = 0 and not IsNull(v) then Description_ = CStr(v)
	End Property  

	'######################################################## 
	'# Date when the table was created. Read-only
	Public Property Get DateCreated
		DateCreated = DateCreated_
	End Property  

	Public Property Let DateCreated(v)
		if IsNull(DateCreated_) and not IsNull(v) then DateCreated_ = CDate(v)
	End Property  

	'######################################################## 
	'# Date when the table was last modified. Read-only
	Public Property Get DateModified
		DateModified = DateModified_
	End Property 
	 
	Public Property Let DateModified(v)
		if IsNull(DateModified_) and not IsNull(v) then DateModified_ = CDate(v)
	End Property  

	'######################################################## 
	'# Creates and appends a new field
	Public Function CreateField(ByRef NewFld, Indexed)
		CreateField = False
		if not DBAE_DEBUG then On Error Resume Next
		
		dim xCat, fld, isUnique, sSQL
		set xCat = Server.CreateObject("ADOX.Catalog")
		If IsEmpty(xCat) or xCat Is Nothing Then
			'ADOX is not available, then let's create the field with pure SQL
			sSQL = "ALTER TABLE [" & Name_ & "] ADD COLUMN " & NewFld.SQL
			call Parent_.JetConnection.Execute(sSQL, adExecuteNoRecords)
		Else
			'whoala! ADOX with us, easy work :)
			set xCat.ActiveConnection = Parent_.JetConnection
			set fld = Server.CreateObject("ADOX.Column")
			set fld.ParentCatalog = xCat
			fld.Name = NewFld.Name
			if NewFld.MaxLength > 0 then fld.DefinedSize = NewFld.MaxLength
			fld.Type = NewFld.FieldType
			fld.Properties("Nullable").Value = NewFld.IsNullable
			if NewFld.IsAutonumber then fld.Properties("Autoincrement").Value = True
			fld.Properties("Jet OLEDB:Compressed UNICODE Strings").Value = NewFld.Compressed
			fld.Properties("Jet OLEDB:Allow Zero Length").Value = NewFld.AllowZeroLength
			if not IsNull(NewFld.Description) then fld.Properties("Description").Value = NewFld.Description
			
			'Do not use Default property. It is not always working
			'if not IsNull(NewFld.DefaultValue) then fld.Properties("Default").Value = NewFld.DefaultValue
			
			xCat.Tables(Name_).Columns.Append fld
			CreateField = not Parent_.IsError
			set fld = nothing
			set xCat = nothing
		End If
		
		if not Parent_.HasError and not IsNull(NewFld.DefaultValue) then
			call Parent_.JetConnection.Execute("ALTER TABLE [" & Name_ & "] ALTER COLUMN [" & NewFld.Name & "] SET DEFAULT " & NewFld.DefaultValue)
		end if
		
		if not Parent_.HasError and Indexed > 0 then
			Randomize
			if Indexed = 2 then isUnique = True else isUnique = False
			CreateIndex "Index_" & CLng(Rnd() * 1000000), NewFld.Name, isUnique, False
		end if
		
	End Function

	'######################################################## 
	'# Deletes an existing field
	Public Function DeleteField(FieldName)
		dim key, sSQL
		
		'find and delete index first
		for each key in Indexes.Keys
			if Indexes_.Item(key).TargetField = CStr(FieldName) then DeleteIndex Indexes_.Item(key).Name, FieldName
		next
		
		'delete the field itself now
		sSQL = "ALTER TABLE [" & Name_ & "] DROP COLUMN [" & FieldName & "]"
		Parent_.JetConnection.Execute sSQL, adExecuteNoRecords
		DeleteField = not Parent_.IsError
		if not Parent_.HasError and Fields_.Exists(FieldName) then Fields_.Remove FieldName
	End Function

	'######################################################## 
	'# Creates a new index
	Public Function CreateIndex(IndexName, TargetField, IsUnique, IsPrimary)
		dim key, str, strPIndex, sSQL
		
		if IsPrimary then
			'save all primary keys first, then delete them
			str = ""
			if Len(IndexName) = 0 then IndexName = "PrimaryKey"
			for each key in Indexes.Keys
				if Indexes_.Item(key).IsPrimary then 
					str = str & "[" & Indexes_.Item(key).TargetField & "],"
					strPIndex = Indexes_.Item(key).Name
				end if
			next
			If Len(str) > 0 then 
				sSQL = "DROP INDEX [" & strPIndex & "] ON [" & Name_ & "]"
				Parent_.JetConnection.Execute sSQL, adExecuteNoRecords
			end if
			sSQL = "CREATE INDEX [" & IndexName & "] ON [" & Name_ & "](" & str & "[" & TargetField & "]) WITH PRIMARY"
			Parent_.JetConnection.Execute sSQL, adExecuteNoRecords
		else
			sSQL = "CREATE "
			if IsUnique then sSQL = sSQL & "UNIQUE "
			sSQL = sSQL & "INDEX [" & IndexName & "] ON [" & Name_ & "]([" & TargetField & "])"
			Parent_.JetConnection.Execute sSQL, adExecuteNoRecords
		end if
		CreateIndex = not Parent_.IsError
		if not Parent_.HasError then
			Indexes_.Item(".uninitialized") = null
		end if
	End Function

	'######################################################## 
	'# Deletes an existing index
	Public Function DeleteIndex(IndexName, FieldName)
		dim key, sSQL, str
		DeleteIndex = False
		
		'find out if this index is primary
		str = ""
		If Indexes.Exists(IndexName & "." & FieldName) then
			if Indexes_.Item(IndexName & "." & FieldName).IsPrimary then
				'save other primary indexes if any
				for each key in Indexes_.Keys
					if key <> IndexName & "." & FieldName and Indexes_.Item(key).IsPrimary then str = str & "[" & Indexes_.Item(key).TargetField & "],"
				next
			end if
			sSQL = "DROP INDEX [" & IndexName & "] ON [" & Name_ & "]"
			Parent_.JetConnection.Execute sSQL, adExecuteNoRecords
			if not Parent_.IsError then
				if Len(str) > 0 then
					're-create all primary keys
					str = Left(str, Len(str) - 1)
					sSQL = "CREATE INDEX [" & IndexName & "] ON [" & Name_ & "](" & str & ") WITH PRIMARY"
					Parent_.JetConnection.Execute sSQL, adExecuteNoRecords
					DeleteIndex = not Parent_.IsError
				end if
				If not Parent_.HasError then Indexes_.Remove IndexName & "." & FieldName
			end if
		end if
	End Function

	'######################################################## 
	'# Returns an ADO Recordset object with data of the table
	Public Function GetRawData(PageSize, Filter, ReadOnly)
		dim rec, lockType
		if Len(Filter) = 0 then Filter = "SELECT * FROM [" & Name_ & "]"
		set rec = Server.CreateObject("ADODB.Recordset")
		if IsNumeric(PageSize) then 
			rec.CacheSize = CInt(PageSize)
			rec.PageSize = CInt(PageSize)
		end if
		if ReadOnly then lockType  = adLockReadOnly else lockType = adLockOptimistic
		rec.Open Filter, Parent_.JetConnection, adOpenKeyset, lockType
		Set GetRawData = rec
	End Function
 
	'######################################################## 
	'# Returns True is the object has been initialized
	Public Function IsInitialized
		if Len(Name_) > 0 and IsObject(Parent_) then IsInitialized = True else IsInitialized = False
	End Function

	'---------------------------
	'protected and private

	Private	Parent_
	Private Indexes_
	Private Fields_
	Private Name_
	Private Description_
	Private DateCreated_
	Private DateModified_

End Class
' END CLASS DEFINITION DBATable


'///////////////////////////////////////////////////
'// ' Class that describes View in database
'//
Class DBAView

	'constructor
	Private Sub Class_Initialize
		Set Parent_		= Nothing
		Name_			= ""
		Body_			= ""
		DateCreated_	= null
		DateModified_	= null
		Description_	= ""
	End Sub
	
	'destructor
	Private Sub Class_Terminate
		Set Parent_ = Nothing
	End Sub


	'######################################################## 
	'# 
	Public Property Set Parent(v)
		Set Parent_ = v
	End Property    
	
	Public Property Get Parent
		Set Parent = Parent_
	End Property  

	'######################################################## 
	'# Name of the View
	Public Property Let Name(v)
		if IsInitialized and Name_ <> v then
			'we are updating the view. Actually just deleting it and creating again
			dim con, sSQL
			sSQL = "DROP VIEW [" & Name_ & "]"
			set con = Parent_.JetConnection
			con.BeginTrans
			if not DBAE_DEBUG then On Error Resume Next
			con.Execute sSQL, adExecuteNoRecords
			call Parent_.IsError
			Name_ = CStr(v)
			con.Execute SQL, adExecuteNoRecords
			if Parent_.IsError then
				con.RollbackTrans
			else
				con.CommitTrans
			end if
		end if
		Name_ = CStr(v)
	End Property    
	
	Public Property Get Name
		Name = Name_
	End Property  

	'######################################################## 
	'# Code of the view
	Public Property Let Body(v)
		if IsInitialized and Body_ <> v then
			dim xCatalog, Command
			if not DBAE_DEBUG then On Error Resume Next
			set xCatalog = Server.CreateObject("ADOX.Catalog")
			If IsEmpty(xCatalog) or xCatalog Is Nothing Then
				'when ADOX is not available. Just re-create the view
				dim con, sSQL
				sSQL = "DROP VIEW [" & Name_ & "]"
				set con = Parent_.JetConnection
				con.BeginTrans
				call con.Execute(sSQL, adExecuteNoRecords)
				call Parent_.IsError
				Body_ = CStr(v)
				con.Execute SQL, adExecuteNoRecords
				if Parent_.IsError then
					con.RollbackTrans
				else
					con.CommitTrans
				end if
			Else
				set Command = Server.CreateObject("ADODB.Command")
				set xCatalog.ActiveConnection = Parent_.JetConnection
				with Command
					.ActiveConnection = Parent_.JetConnection
					.CommandText = CStr(v)
					.CommandType = adCmdText
				end with
				
				set xCatalog.Views(Name_).Command = Command
				if not Parent_.IsError then Body_ = CStr(v)
				
				set Command = Nothing
				set xCatalog = Nothing
			End If
		end if
		Body_ = CStr(v)
	End Property    
	
	Public Property Get Body
		Body = Body_
	End Property  

	'######################################################## 
	'# SQL statement that can be used to create this view
	Public Property Get SQL
		'PROCEDURE instead of VIEW just to avoid "Only simple queries.." error
		SQL = "CREATE PROCEDURE [" & Name_ & "] AS " & vbCrLf & Body_
	End Property  

	'######################################################## 
	'# 
	Public Property Let Description(v)
		if Len(Description_) = 0 and not IsNull(v) then Description_ = CStr(v)
	End Property    
	
	Public Property Get Description
		Description = Description_
	End Property  

	'######################################################## 
	'# 
	Public Property Let DateCreated(v)
		if IsDate(v) and IsNull(DateCreated_) then DateCreated_ = CDate(v)
	End Property    
	
	Public Property Get DateCreated
		DateCreated = DateCreated_
	End Property  

	'######################################################## 
	'# 
	Public Property Let DateModified(v)
		if IsDate(v) and IsNull(DateModified_) then DateModified_ = CDate(v)
	End Property    
	
	Public Property Get DateModified
		DateModified = DateModified_
	End Property  

	'######################################################## 
	'# Returns True if the object has been initizliazed
	Public Function IsInitialized
		if Len(Name_) > 0 and Len(Body_) > 0 then IsInitialized = True else IsInitialized = False
	End Function

	'---------------------------
	'protected and private
	Private Parent_
	Private Name_
	Private Body_
	Private Description_
	Private DateCreated_
	Private DateModified_

End Class
' END CLASS DEFINITION DBAView


'///////////////////////////////////////////////////
'// ' Class that describes a single Procedure in database
'//
Class DBAProcedure



	'######################################################## 
	'# 
	Public Property Set Parent(v)
		Set Parent_ = v
	End Property    
	
	Public Property Get Parent
		Set Parent = Parent_
	End Property  

	'######################################################## 
	'# Name of procedure
	Public Property Let Name(v)
		if IsInitialized and Name_ <> v then
			'we are updating the procedure. Actually just deleting it and creating again
			dim con, sSQL
			sSQL = "DROP PROCEDURE [" & Name_ & "]"
			set con = Parent_.JetConnection
			con.BeginTrans
			if not DBAE_DEBUG then On Error Resume Next
			con.Execute sSQL, adExecuteNoRecords
			call Parent_.IsError
			Name_ = CStr(v)
			con.Execute SQL, adExecuteNoRecords
			if Parent_.IsError then
				con.RollbackTrans
			else
				con.CommitTrans
			end if
		end if
		Name_ = CStr(v)
	End Property    
	
	Public Property Get Name
		Name = Name_
	End Property  

	'######################################################## 
	'# Procedure's code
	Public Property Let Body(v)
		if IsInitialized and Body_ <> v then
			dim xCatalog, Command
			if not DBAE_DEBUG then On Error Resume Next
			set xCatalog = Server.CreateObject("ADOX.Catalog")
			If IsEmpty(xCatalog) or xCatalog Is Nothing Then
				'when ADOX is not available. Just re-create the view
				dim con, sSQL
				sSQL = "DROP PROCEDURE [" & Name_ & "]"
				set con = Parent_.JetConnection
				con.BeginTrans
				call con.Execute(sSQL, adExecuteNoRecords)
				call Parent_.IsError
				Body_ = CStr(v)
				con.Execute SQL, adExecuteNoRecords
				if Parent_.IsError then
					con.RollbackTrans
				else
					con.CommitTrans
				end if
			Else
				set Command = Server.CreateObject("ADODB.Command")
				set xCatalog.ActiveConnection = Parent_.JetConnection
				with Command
					.ActiveConnection = Parent_.JetConnection
					.CommandText = CStr(v)
					.CommandType = adCmdText
				end with
				
				set xCatalog.Procedures(Name_).Command = Command
				if not Parent_.IsError then Body_ = Command.CommandText
				
				set Command = Nothing
				set xCatalog = Nothing
			End If
		end if
		Body_ = CStr(v)
	End Property    
	
	Public Property Get Body
		Body = Body_
	End Property  

	'######################################################## 
	'# SQL statement needed to create such procedure
	Public Property Get SQL
		SQL =	"CREATE PROCEDURE [" & Name_ & "] AS " & vbCrLf & Body_
	End Property  

	'######################################################## 
	'# Description of procedure (read-only)
	Public Property Let Description(v)
		if Len(Description_) = 0 and not IsNull(v) then Description_ = CStr(v)
	End Property    
	
	Public Property Get Description
		Description = Description_
	End Property  

	'######################################################## 
	'# Date when the procedure was created. Read-only
	Public Property Let DateCreated(v)
		if IsDate(v) and IsNull(DateCreated_) then DateCreated_ = CDate(v)
	End Property    
	
	Public Property Get DateCreated
		DateCreated = DateCreated_
	End Property  

	'######################################################## 
	'# Date when the procedure was last modified. Usually same as DateCreated. Read-only
	Public Property Let DateModified(v)
		if IsDate(v) and IsNull(DateModified_) then DateModified_ = CDate(v)
	End Property    
	
	Public Property Get DateModified
		DateModified = DateModified_
	End Property  


	'######################################################## 
	'# Returns True is the object has been properly initialized
	Public Function IsInitialized()
		if Len(Name_) > 0 and Len(Body_) > 0 then IsInitialized = True else IsInitialized = False
	End Function
 

	'---------------------------
	'protected and private
	Private Parent_
	Private Name_
	Private Body_
	Private Description_
	Private DateCreated_
	Private DateModified_


	' Constructor
	Private Sub Class_Initialize()
		Set Parent_		= Nothing
		Name_			= ""
		Body_			= ""
		DateCreated_	= null
		DateModified_	= null
		Description_	= ""
	End Sub
	
	' Destructor
	Private Sub Class_Terminate()
		Set Parent_ = Nothing
	End Sub


End Class
' END CLASS DEFINITION DBAProcedure


'///////////////////////////////////////////////////
'// ' Class describes single field in a table
'//
Class DBAField



	'######################################################## 
	'# 
	Public Property Set Parent(v)
		Set Parent_ = v
	End Property    
	
	Public Property Get Parent
		Parent = Parent_
	End Property  

	'######################################################## 
	'# 
	Public Property Let Name(v)
		if Len(Name_) > 0 then 
			'change the name of the column
			dim xCat
			set xCat = Server.CreateObject("ADOX.Catalog")
			If IsEmpty(xCat) or xCat Is Nothing Then
				Parent_.Parent.LastError = "ADOX is not available. Couldn't change column's name"
				v = Name_
			Else
				set xCat.ActiveConnection = Parent_.Parent.JetConnection
				xCat.Tables(Parent_.Name).Columns(Name_).Name = CStr(v)
				set xCat = Nothing
			End If
		end if
		Name_ = CStr(v)
	End Property    
	
	Public Property Get Name
		Name = Name_
	End Property  

	'######################################################## 
	'# sets/returns field type
	Public Property Let FieldType(v)
		If FieldType_ >= 0 and v <> FieldType_ then PendingUpdates_ = True
		if IsNumeric(v) then 
			FieldType_ = CLng(v)
		else
			Select Case UCase(v)
				Case "COUNTER"			IsAutonumber_ = True : FieldType_ = 3
				Case "LONG"				FieldType_ = 3
				Case "DATETIME"			FieldType_ = 7
				Case "BIT"				FieldType_ = 11
				Case "MONEY"			FieldType_ = 6
				Case "BINARY"			FieldType_ = 128
				Case "TINYINT"			FieldType_ = 17
				Case "DECIMAL"			FieldType_ = 131
				Case "FLOAT"			FieldType_ = 5
				Case "INTEGER"			FieldType_ = 2
				Case "REAL"				FieldType_ = 4
				Case "UNIQUEIDENTIFIER"	FieldType_ = 72
				Case "MEMO"				MaxLength_ = 0 : FieldType_ = 203
				Case "TEXT"				FieldType_ = 130
				Case Else				FieldType_ = -1
			End Select
		end if
	End Property    
	
	Public Property Get FieldType
		call UpdateBatch
		
		FieldType = FieldType_
	End Property  

	'######################################################## 
	'# 
	Public Property Let MaxLength(v)
		if not IsEmpty(MaxLength_) and v <> MaxLength_ then PendingUpdates_ = True
		if IsNumeric(v) then MaxLength_ = CInt(v) else MaxLength_ = -1
	End Property    
	
	Public Property Get MaxLength
		call UpdateBatch
		
		MaxLength = MaxLength_
	End Property  

	'######################################################## 
	'# 
	Public Property Get IsPrimaryKey
		if IsNull(IsPrimaryKey_) then
			dim key
			IsPrimaryKey_ = False
			for each key in Parent_.Indexes.Keys
				if Parent_.Indexes.Item(key).TargetField = Name_ and Parent_.Indexes.Item(key).IsPrimary then
					IsPrimaryKey_ = True
					Exit for
				end if
			next
		end if
		
		IsPrimaryKey = IsPrimaryKey_
	End Property  

	'######################################################## 
	'# 
	Public Property Let IsAutonumber(v)
		if not IsEmpty(IsAutonumber_) and not IsNull(v) and v <> IsAutonumber_ then PendingUpdates_ = True
		if not IsNull(v) then IsAutonumber_ = CBool(v)
	End Property
	
	Public Property Get IsAutonumber
		IsAutonumber = IsAutonumber_
	End Property  

	'######################################################## 
	'# 
	Public Property Let Ordinal(v)
		if Ordinal_ = 0 then Ordinal_ = CInt(v)
	End Property    
	
	Public Property Get Ordinal
		Ordinal = Ordinal_
	End Property  

	'######################################################## 
	'# 
	Public Property Get HasDefault
		HasDefault = not IsNull(DefaultValue_) and not IsEmpty(DefaultValue_)
	End Property  

	'######################################################## 
	'# 
	Public Property Let DefaultValue(v)
		if not IsEmpty(DefaultValue_) and v <> DefaultValue_ then PendingUpdates_ = True
		DefaultValue_ = v
	End Property    
	
	Public Property Get DefaultValue
		call UpdateBatch
		
		DefaultValue = DefaultValue_
	End Property  

	'######################################################## 
	'# 
	Public Property Let IsNullable(v)
		if not IsEmpty(IsNullable_) and v <> IsNullable_ then PendingUpdates_ = True
		IsNullable_ = CBool(v)
	End Property    
	
	Public Property Get IsNullable
		IsNullable = IsNullable_
	End Property  

	'######################################################## 
	'# 
	Public Property Let Description(v)
		if not IsNull(v) and v <> Description_ and not IsEmpty(Description_) then PendingUpdates_ = True
		if IsNull(v) then Description_ = "" else Description_ = CStr(v)
	End Property    
	
	Public Property Get Description
		Description = Description_
	End Property  

	'######################################################## 
	'# 
	Public Property Let AllowZeroLength(v)
		if not IsEmpty(AllowZeroLength_) and not IsNull(v) and v <> AllowZeroLength_ then PendingUpdates_ = True
		AllowZeroLength_ = CBool(v)
	End Property
	
	Public Property Get AllowZeroLength
		AllowZeroLength = AllowZeroLength_
	End Property

	'######################################################## 
	'# 
	Public Property Let Compressed(v)
		if not IsNull(v) then Compressed_ = CBool(v)
	End Property

	Public Property Get Compressed
		Compressed = Compressed_
	End Property

	'######################################################## 
	'# return SQL string for this field
	Public Property Get SQL
		call UpdateBatch
		
		dim strSQL
		strSQL = "[" & Name_ & "] " & GetSQLTypeName()
		if GetSQLTypeName() = "TEXT" then strSQL = strSQL & "(" & MaxLength_ & ")"
		if not IsNullable_ then strSQL = strSQL & " NOT NULL"
		if HasDefault then strSQL = strSQL & " DEFAULT " & DefaultValue_
		SQL = strSQL
	End Property  

	'######################################################## 
	'# 	
	Public Function IsInitialized()
		if Len(Name_) > 0 and FieldType_ >= 0 and TypeName(Parent_) <> "Nothing" then IsInitialized = True else IsInitialized = False
	End Function
	
	'######################################################## 
	'# Returns SQL type name
	Function GetSQLTypeName
		Select Case FieldType_
		Case 3		if IsAutonumber then GetSQLTypeName = "COUNTER" else GetSQLTypeName = "LONG"
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
		Case 130	if MaxLength_ = 0 then GetSQLTypeName = "MEMO" else GetSQLTypeName = "TEXT"
		Case 202	GetSQLTypeName = "TEXT"
		Case 203	GetSQLTypeName = "MEMO"
		Case Else	GetSQLTypeName = ""
		End Select
	End Function
 
	'######################################################## 
	'# Returns human-readable name of the type, as it is in Access	
	Function GetTypeName
		Select Case FieldType_
		Case 3		if IsAutonumber then GetTypeName = "AutoNumber" else GetTypeName = "Long Integer"
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
		Case 130	if MaxLength_ = 0 then GetTypeName = "Memo" else GetTypeName = "Text"
		Case 202	GetTypeName = "Text"
		Case 203	GetTypeName = "Memo"
		Case Else	GetTypeName = ""
		End Select
	End Function
	
	'######################################################## 
	'# Updates any changes made to the field. Triggered from almost all functions and properties	
	Function UpdateBatch
		if not PendingUpdates_ or TypeName(Parent_) = "Nothing" then 
			UpdateBatch = True
			Exit Function
		end if

		dim xCat, field, sSQL, sSQLType
		
		if not DBAE_DEBUG then On Error Resume Next
		sSQLType = GetSQLTypeName
		sSQL = "ALTER TABLE [" & Parent_.Name & "] ALTER COLUMN [" & Name_ & "] " & sSQLType
		if sSQLType = "TEXT" then sSQL = sSQL & "(" & MaxLength_ & ")"
		if not IsNullable then sSQL = sSQL & " NOT NULL"
		Parent_.Parent.JetConnection.Execute sSQL, adExecuteNoRecords
		if not Parent_.Parent.IsError then
			'set other field properties
			set xCat = Server.CreateObject("ADOX.Catalog")
			if not IsEmpty(xCat) and not xCat Is Nothing Then
				set xCat.ActiveConnection = Parent_.Parent.JetConnection
				set field = xCat.Tables(Parent_.Name).Columns(Name_)
				with field
					if sSQLType = "TEXT" or sSQLType = "MEMO" then
						.Properties("Jet OLEDB:Allow Zero Length").Value = AllowZeroLength_
					end if
					if not IsNull(DefaultValue_) then .Properties("Default").Value = DefaultValue_
					if not IsNull(Description_) then .Properties("Description").Value = Description_
				end with
				set field = Nothing
				set xCat = Nothing
				Parent_.Parent.IsError
			End If
		end if
		
		UpdateBatch = not Parent_.Parent.HasError
		PendingUpdates_ = False
		
		'if error occured, let parent reload fields
		if Parent_.Parent.HasError then Parent_.Fields.Item(".uninitialized") = null
	End Function

	'---------------------------
	'protected and private

	Private Parent_
	Private Name_
	Private FieldType_
	Private MaxLength_
	Private IsPrimaryKey_
	Private IsAutonumber_
	Private Ordinal_
	Private DefaultValue_
	Private IsNullable_
	Private Description_
	Private PendingUpdates_
	Private AllowZeroLength_
	Private Compressed_


	' Constructor
	Private Sub Class_Initialize()
		Set Parent_		= Nothing
		Name_			= ""
		FieldType_		= -1
		MaxLength_		= Empty
		IsPrimaryKey_	= null
		IsAutonumber_	= Empty
		Ordinal_		= 0
		DefaultValue_	= Empty
		IsNullable_		= Empty
		Description_	= Empty
		PendingUpdates_	= False
		AllowZeroLength_= Empty
		Compressed_		= Empty
	End Sub

	' Destructor
	Private Sub Class_Terminate()
		call UpdateBatch
		
		Set Parent_ = Nothing
	End Sub


End Class
' END CLASS DEFINITION DBAField



'///////////////////////////////////////////////////
'// ' Holds information about particular index in the table
'//
Class DBAIndex



	'######################################################## 
	'# 
	Public Property Set Parent(v)
		Set Parent_ = v
	End Property    
	
	Public Property Get Parent
		Parent = Parent_
	End Property  

	'######################################################## 
	'# 
	Public Property Let Name(v)
		Name_ = CStr(v)
	End Property    
	
	Public Property Get Name
		Name = Name_
	End Property  

	'######################################################## 
	'# 
	Public Property Let TargetField(v)
		TargetField_ = CStr(v)
	End Property    
	
	Public Property Get TargetField
		TargetField = TargetField_
	End Property  

	'######################################################## 
	'# 
	Public Property Let IsUnique(v)
		if not IsNull(v) then IsUnique_ = CBool(v)
	End Property    
	
	Public Property Get IsUnique
		IsUnique = IsUnique_
	End Property  

	'######################################################## 
	'# 
	Public Property Let IsPrimary(v)
		if not IsNull(v) then IsPrimary_ = CBool(v)
	End Property    
	
	Public Property Get IsPrimary
		IsPrimary = IsPrimary_
	End Property  

	'######################################################## 
	'# Returns True is the index is actually a foreign key
	Public Property Get IsForeignKey
		if IsNull(IsForeignKey_) then
			dim rec
			IsForeignKey_ = False
			set rec = Parent_.Parent.JetConnection.OpenSchema(adSchemaForeignKeys, Array(empty, empty, empty, empty, empty, Parent_.Name))
			do while not rec.EOF
				if rec("FK_NAME") = Name_ then
					IsForeignKey_ = True
					Exit Do
				end if
				rec.MoveNext
			loop
			rec.close
			set rec = nothing
		end if
		
		IsForeignKey = IsForeignKey_
	End Property  

	'######################################################## 
	'# returns SQL statement that describes this index
	Public Property Get SQL
		dim strSQL, item
		strSQL = "CREATE "
		if IsUnique_ and not IsPrimary_ then strSQL = strSQL & "UNIQUE "
		
		strSQL = strSQL & "INDEX [" & Name_ & "] ON [" & Parent_.Name & "]("
		'go through all indexes in the table to find same index to different field
		for each item in Parent_.Indexes.Items
			if item.Name = Name_ and item.TargetField <> TargetField_ then strSQL = strSQL & "[" & item.TargetField & "],"
		next
		strSQL = strSQL & "[" & TargetField_ & "])"
		if IsPrimary_ then strSQL = strSQL & " WITH PRIMARY"
		
		SQL = strSQL
	End Property  

	'######################################################## 
	'# 	
	Public Function IsInitialized()
		if IsObject(Parent_) and Len(Name_) > 0 and Len(TargetField_) > 0 then IsInitialized = True else IsInitialized = False
	End Function


	'---------------------------
	'protected and private
	Private Parent_
	Private Name_
	Private TargetField_
	Private IsUnique_
	Private IsPrimary_
	Private IsForeignKey_

	'######################################################## 
	'# Constructor
	Private Sub Class_Initialize()
		Set Parent_		= Nothing
		Name_			= ""
		TargetField_	= ""
		IsUnique_		= False
		IsPrimary_		= False
		IsForeignKey_	= null
	End Sub

	'######################################################## 
	'# Destructor
	Private Sub Class_Terminate()
		Set Parent_ = Nothing
	End Sub


End Class
' END CLASS DEFINITION DBAIndex



'///////////////////////////////////////////////////
'// ' Class that describes a single relatio between 2 tables
'//
Class DBARelation



	'######################################################## 
	'# 
	Public Property Set Parent(v)
		Set Parent_ = v
	End Property    
	
	Public Property Get Parent
		Parent = Parent_
	End Property  

	'######################################################## 
	'# 
	Public Property Let Name(v)
		if not IsNull(v) then Name_ = CStr(v)
	End Property    
	
	Public Property Get Name
		Name = Name_
	End Property  

	'######################################################## 
	'# 
	Public Property Let PrimaryTable(v)
		PrimaryTable_ = CStr(v)
	End Property    
	
	Public Property Get PrimaryTable
		PrimaryTable = PrimaryTable_
	End Property  

	'######################################################## 
	'# 
	Public Property Let PrimaryField(v)
		PrimaryField_ = CStr(v)
	End Property    
	
	Public Property Get PrimaryField
		PrimaryField = PrimaryField_
	End Property  

	'######################################################## 
	'# 
	Public Property Let PrimaryIndex(v)
		if not IsNull(v) then PrimaryIndex_ = CStr(v)
	End Property    
	
	Public Property Get PrimaryIndex
		PrimaryIndex = PrimaryIndex_
	End Property  

	'######################################################## 
	'# 
	Public Property Let ForeignTable(v)
		ForeignTable_ = CStr(v)
	End Property    
	
	Public Property Get ForeignTable
		ForeignTable = ForeignTable_
	End Property  

	'######################################################## 
	'# 
	Public Property Let ForeignField(v)
		ForeignField_ = CStr(v)
	End Property    
	
	Public Property Get ForeignField
		ForeignField = ForeignField_
	End Property  

	'######################################################## 
	'# 
	Public Property Let OnUpdate(v)
		if not IsNull(v) then OnUpdate_ = CStr(v)
	End Property    
	
	Public Property Get OnUpdate
		if IsNull(OnUpdate_) then OnUpdate = "NO ACTION" else OnUpdate = OnUpdate_
	End Property  

	'######################################################## 
	'# 
	Public Property Let OnDelete(v)
		if not IsNull(v) then OnDelete_ = CStr(v)
	End Property    
	
	Public Property Get OnDelete
		if IsNull(OnDelete_) then OnDelete = "NO ACTION" else OnDelete = OnDelete_
	End Property  

	'######################################################## 
	'# returns SQL statement that describes the relation
	Public Property Get SQL
		if not IsInitialized then Exit Property
		
		dim strSQL
		strSQL =	"ALTER TABLE [" & ForeignTable_ & "] ADD CONSTRAINT [" & Name_ & "] " &_
					"FOREIGN KEY ([" & ForeignField_ & "]) " & vbCrLf &_
					"REFERENCES [" & PrimaryTable_ & "] ([" & PrimaryField_ & "])"
		if Len(OnUpdate_) > 0 then strSQL = strSQL & " ON UPDATE " & OnUpdate_
		if Len(OnDelete_) > 0 then strSQL = strSQL & " ON DELETE " & OnDelete_
		strSQL = strSQL & ";"
		
		SQL = strSQL
	End Property  


	'######################################################## 
	'# Returns True if the object has been initialized
	Public Function IsInitialized()
		if 	Len(PrimaryTable_) > 0 and _
			Len(PrimaryField_) > 0 and _
			Len(ForeignTable_) > 0 and _
			Len(ForeignField_) > 0 and _
			IsObject(Parent_) _
			then IsInitialized = True else IsInitialized = False
	End Function
 

	'---------------------------
	'protected and private
	Private Parent_
	Private Name_
	Private PrimaryTable_
	Private PrimaryField_
	Private PrimaryIndex_
	Private ForeignTable_
	Private ForeignField_
	Private OnDelete_
	Private OnUpdate_

	'######################################################## 
	'# Constructor
	Private Sub Class_Initialize()
		Set Parent_		= Nothing
		PrimaryTable_	= ""
		PrimaryField_	= ""
		PrimaryIndex_	= ""
		ForeignTable_	= ""
		ForeignField_	= ""
		OnDelete_		= ""
		OnUpdate_		= ""
	End Sub

	'######################################################## 
	'# Destructor
	Private Sub Class_Terminate()
		Set Parent_		= Nothing
	End Sub


End Class
' END CLASS DEFINITION DBARelation
%>

