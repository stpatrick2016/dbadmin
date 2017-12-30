<SCRIPT LANGUAGE=vbscript RUNAT=Server>

Class TableIndexes
	Private sTableName, sPKColumns, sPKIndexName, sIXColumns, sIndexes, sForeignKeys,Separator
	

	'#Opens or refreshes indexes list
	Public Sub OpenTable(sTable)
		Separator = vbTab
		sTableName = sTable
		sPKColumns = ""
		sPKIndexName = ""
		sIXColumns = ""
		sIndexes = ""
		sForeignKeys = ""

		dim con, rec
		OpenConnection con
		
		'get all indexes
		set rec = con.OpenSchema(adSchemaIndexes,Array(empty,empty,empty,empty, sTableName))
		do while not rec.EOF and Err=0
			if rec("PRIMARY_KEY") then AppendPrimaryKey rec("COLUMN_NAME"), rec("INDEX_NAME")
			AppendIndex rec("COLUMN_NAME"), rec("INDEX_NAME"), rec("UNIQUE")
			rec.MoveNext
		loop
		rec.close
		
		'get foreign keys
		set rec = con.OpenSchema(adSchemaForeignKeys, Array(empty,empty,empty,empty,empty,sTableName))
		do while not rec.EOF and Err=0
			sForeignKeys = sForeignKeys & rec("FK_NAME") & Separator
			rec.MoveNext
		loop
		rec.Close
		
		con.Close
		set rec = nothing
		set con = nothing
	End Sub

	'#Determines if column is a primary key
	Function IsPrimaryKey(column)
		if InStr(1, sPKColumns, column & Separator, vbTextCompare) > 0 then
			IsPrimaryKey = True
		else
			IsPrimaryKey = False
		end if
	End Function
	
	'#checks if the index is a foreign key
	Function IsForeignKey(index)
		if InStr(1, sForeignKeys, index & Separator, vbTextCompare) > 0 then
			IsForeignKey = True
		else
			IsForeignKey = False
		end if
	End Function

	'#Returns an array of Primary Keys
	Function GetPrimaryKeys
		GetPrimaryKeys = Split(sPKColumns, Separator)
	End Function
	
	'#Returns an array with all indexes for this table.
	'#Structure of array: (index)(index_name,column_name,unique)
	Function GetIndexes
		dim ar1, ret(), ar2, ar3, i
		ar1 = Split(sIndexes, Separator)
		ar3 = Split(sIXColumns, Separator)
		Redim ret(UBound(ar1))

		for i=0 to UBound(ar1)
			if Len(ar1(i)) > 0 then
				ar2 = Split(ar1(i), "!")
				ret(i) = Array(ar2(0), ar3(i), ar2(1))
			end if
		next
		
		GetIndexes = ret
	End Function
	
	'#Returns name of Primary Key index
	Function GetPrimaryKeyName
		GetPrimaryKeyName = sPKIndexName
	End Function
	
	'#create a new index
	Function CreateIndex(index,column,primary,unique)
		dim con, sSQL, s
		OpenConnection con
		if primary then
			if GetPrimaryKeysCount() > 0 then
				sSQL = "DROP INDEX [" & index & "] ON [" & sTableName & "]"
				con.Execute sSQL, adExecuteNoRecords
			end if
			AppendPrimaryKey column, index
			dim ar
			ar = GetPrimaryKeys()
			sSQL = "CREATE INDEX [" & sPKIndexName & "] ON [" & sTableName & "] ("
			for each s in ar
				if Len(s) > 0 then sSQL = sSQL & "[" & s & "],"
			next
			if Right(sSQL, 1) = "," then sSQL = Left(sSQL, Len(sSQL)-1)
			sSQL = sSQL & ") WITH PRIMARY"
			unique = True
		else
			sSQL = "CREATE "
			if unique then sSQL = sSQL & "UNIQUE "
			sSQL = sSQL & "INDEX [" & index & "] ON [" & sTableName & "] ([" & column & "])"
		end if

		con.Execute sSQL, adExecuteNoRecords
		If Err then 
			CreateIndex = Err.Description
		else
			CreateIndex = ""
			AppendIndex column, index, unique
		end if
		con.close
		set con = nothing
	End Function
	
	'#deletes index and/or primary key
	Function DeleteIndex(index, column)
		dim con, sSQL, s
		OpenConnection con
		sSQL = "DROP INDEX [" & index & "] ON [" & sTableName & "]"
		con.Execute sSQL, adExecuteNoRecords
		if Err then
			DeleteIndex = err.Description 
			con.close
			set con = nothing
			Exit Function
		end if
		
		if index = sPKIndexName then	
			if GetPrimaryKeysCount() > 1 then
				dim arKeys
				sPKColumns = Replace(sPKColumns, column & Separator, "", 1, -1, vbTextCompare)
				arKeys = GetPrimaryKeys()
				sSQL = "CREATE INDEX [" & index & "] ON [" & sTableName & "] ("
				for each s in arKeys
					if Len(s) > 0 then sSQL = sSQL & "[" & s & "],"
				next
				if Right(sSQL, 1) = "," then sSQL = Left(sSQL, Len(sSQL)-1)
				sSQL = sSQL & ") WITH PRIMARY"
				con.Execute sSQL, adExecuteNoRecords
				if Err then
					DeleteIndex = err.Description 
					con.close
					set con = nothing
					Exit Function
				end if
			else
				sPKIndexName = ""
				sPKColumns = ""
			end if
		end if
		
		'Remove index from indexes array
		dim arIndexes, i
		arIndexes = GetIndexes
		sIXColumns = ""
		sIndexes = ""
		for i=0 to UBound(arIndexes)-1
			if StrComp(arIndexes(i)(0), index, vbTextCompare)<>0 or StrComp(arIndexes(i)(1), column, vbTextCompare)<>0 then
				AppendIndex arIndexes(i)(1), arIndexes(i)(0),arIndexes(i)(2)
			end if
		Next
		con.Close
		set con = nothing
	End Function




	'#Appends primary key to array
	Private Sub AppendPrimaryKey(column,index)
		if InStr(1, sPKColumns, column & Separator, vbTextCompare) <= 0 then	
			sPKColumns = sPKColumns & column & Separator
			sPKIndexName = index
		end if
	End Sub
	
	'#Appends index
	Private Sub AppendIndex(column, index, unique)
		dim arIndexes, i, Found
		arIndexes = GetIndexes
		Found = False
		for i=0 to UBound(arIndexes)-1
			if StrComp(arIndexes(i)(0),index, vbTextCompare)=0 and StrComp(arIndexes(i)(1),column, vbTextCompare)=0 then
				Found = True
				Exit For
			end if
		Next
		if not Found then	
			sIXColumns = sIXColumns & column & Separator
			sIndexes = sIndexes & index & "!" & unique & Separator
		end if
	End Sub
	
	'#Returns number of primary keys defined
	Private Function GetPrimaryKeysCount
		dim ret
		if Len(sPKColumns) = 0 then
			ret = 0
		else
			ret = UBound(Split(sPKColumns, Separator))
		end if
		GetPrimaryKeysCount = ret
	End Function
	
End Class

Function HighlightSQL(sSQL)
	Const KeyWords =	"CREATE|TABLE|COUNTER|NOT NULL|DEFAULT|INDEX|ON|PRIMARY|WITH|LONG|TEXT|DATETIME|BIT|MONEY|BINARY|TINYINT|DECIMAL|FLOAT|INTEGER|REAL|UNIQUEIDENTIFIER|MEMO|UNIQUE|INSERT|INTO|SELECT|FROM|WHERE|UPDATE|DELETE|VALUES|PARAMETERS|ORDER BY|OR|AND|IN|SUM|AS|TOP|SET|LEFT|RIGHT|INNER|JOIN|ASC|DESC|GROUP BY|HAVING|CONSTRAINT|ADD|COLUMN|CASCADE|DROP|TOP|DISTINCT|DISTINCTROW|KEY|MIN|MAX|COUNT|AVG|PROCEDURE|VIEW|STDEV|STDEVP|UNION|ALTER|REFERENCES|FOREIGN"
	
	dim RegEx, s
	set RegEx = new RegExp
	RegEx.Global = True
	RegEx.IgnoreCase = true
	
	'Replace code
	RegEx.Pattern = "(\b" & Replace(KeyWords, "|", "\b|\b") & "\b)"
	sSQL = RegEx.Replace(sSQL, "<font color=""blue"">$1</font>")
	
	'replace numbers
	RegEx.Pattern = "([\s\(<>=\-\+])([0-9]+)([\s,;\)<>=\-\+])"
	sSQL = RegEx.Replace(sSQL, "$1<font color=""green"">$2</font>$3")
	
	set RegEx = nothing
	HighlightSQL = sSQL
End Function

Sub IsError
	if Err then
		Response.Write "<p align=center class=""Error"">" & Err.Description & "</p>"
		Response.Write "</td></tr></table></body></html>"
		Response.End 
	end if
End Sub

Sub OpenConnection(ByRef objCon)
	dim constr
	set objCon = Server.CreateObject("ADODB.Connection")
	constr = strProvider & Session("DBAdminDatabase")
	if Len(CStr(Session("DBAdminDBPassword"))) > 0 then constr = constr & ";Jet OLEDB:Database password=" & CStr(Session("DBAdminDBPassword"))
	objCon.Open constr
End Sub

</SCRIPT>
