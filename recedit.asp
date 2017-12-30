<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href=default.css rel=stylesheet>
<TITLE>Database Administration</TITLE>
</HEAD>
<BODY>
<!--#include file=config.asp -->
<!--#include file=inc_protect.asp -->
<!--#include file=inc_functions.asp -->
<TABLE WIDTH=100% ALIGN=center>
	<TR>
		<TD width=180px valign=top><!--#include file=inc_nav.asp --></TD>
		<TD>
		
		
<%
	Const TableNameKey = "BBC017D1-0A13-4a9d-9A53-52A0CC6A7540"
	Const PKNameKey = "57AFDC29-37C8-48e1-96BE-12D7B79C1825"
	Const ActionNameKey = "714A51CC-797B-4ce5-99C7-81DB8721D68B"
	Const KeyNameKey = "0C61DD31-D805-4f58-A369-E4F33595FB86"
	
'	On error Resume NExt
	dim con, rec, sSQL, sTable, pk, key, fld, bIsEdit, sName, i
	sTable = Request("table")
	pk = Split(Request("pk"), ".")
	key = Split(Request("key"), ".")
	if Request.QueryString("action") = "edit" then 
		bIsEdit = True
		sName = "Update"
	else
		bIsEdit = False
		sName = "Add"
	end if
	set con = Server.CreateObject("ADODB.Connection")
	con.Open strProvider & Session("DBAdminDatabase")
	set rec = Server.CreateObject("ADODB.Recordset")
	IsError
	
	if Request.Form(TableNameKey).Count > 0 then
		sTable = Request.Form(TableNameKey)
		pk = Split(Request.Form(PKNameKey), ".")
		key = Split(Request.Form(KeyNameKey), ".")
		if Request.Form(ActionNameKey) = "edit" then
			sSQL = "SELECT * FROM [" & sTable & "] WHERE"
			fld = ""
			for i=0 to UBound(pk)-1
				if Len(pk(i)) > 0 and Len(key(i)) > 0 then
					sSQL = sSQL & fld & " [" & pk(i) & "]=" & key(i)
					fld = " AND"
				end if
			Next
			bIsEdit = True
		else
			sSQL = "[" & sTable & "]"
			bIsEdit = False
		end if
		rec.Open sSQL, con, adOpenDynamic, adLockOptimistic
		if rec.EOF or bIsEdit = False then 
			rec.AddNew 
			bIsEdit = False
		else
			bIsEdit = True
		end if
		for each fld in rec.Fields 
			if not fld.Properties("ISAUTOINCREMENT") and Len(fld.Name) > 0 then
				if fld.Type = adDate then 
					fld.Value = CDate(Request.Form(fld.Name)) 
				else 
					fld.Value = Request.Form(fld.Name)
				end if
			end if
		Next
		rec.Update 
		if Err then
			Response.Write "<DIV class=Error align=center>" & Err.Description & "</DIV>"
			rec.Close
		else
			rec.Close
			con.Close
			set rec = nothing
			set con = nothing
			Response.Redirect "data.asp?table=" & sTable
		end if
	end if
	
	if Request.QueryString("action") = "edit" then 
		sSQL = "SELECT * FROM [" & sTable & "] WHERE"
		fld = ""
		for i=0 to UBound(pk)-1
			if Len(pk(i)) > 0 and Len(key(i)) > 0 then
				sSQL = sSQL & fld & " [" & pk(i) & "]=" & key(i)
				fld = " AND"
			end if
		Next
	else
		sSQL = "[" & sTable & "]"
	end if
	rec.Open sSQL, con, adOpenForwardOnly, adLockReadOnly 
%>

<H1><%=sName%> a record</H1>
<P align=center>Note that AutoNumber fields you cannot edit as they are updated automatically.
Also columns of type Binary won't be shown here</P>
<FORM action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method=POST>
<INPUT type=hidden name="<%=TableNameKey%>" value="<%=sTable%>">
<INPUT type=hidden name="<%=PKNameKey%>" value="<%=Request.QueryString("pk")%>">
<INPUT type=hidden name="<%=ActionNameKey%>" value="<%=Request.QueryString("action")%>">
<INPUT type=hidden name="<%=KeyNameKey%>" value="<%=Request.QueryString("key")%>">
<TABLE border=0 class=RegularTable align=center>

<%	for each fld in rec.Fields%>
		<%if fld.Type <> adBinary and fld.Type <> adVarBinary and fld.Type <> adLongVarBinary then%>
	<TR>
		<TD style="border:2px outset" valign=top><B><%=fld.Name%></B>&nbsp;(<%=GetTypeName(fld.Type)%>)</TD>
		<TD>
			<%if fld.Type = 203 or fld.Type = 201 then%>
			<TEXTAREA name="<%=fld.Name%>" rows=4 cols=40><%if bIsEdit then Response.Write fld.Value%></TEXTAREA>
			<%elseif fld.Properties("ISAUTOINCREMENT") then%>
			AutoNumber
			<%else%>
			<INPUT type=text name="<%=fld.Name%>" value="<%if bIsEdit then Response.Write fld.Value%>">
			<%end if%>
		</TD>
	</TR>
		<%end if%>
<%	Next%>

</TABLE>
<P align=center>
	<INPUT type=submit value="<%=sName%>" class=button>
	<INPUT type=reset value="Restore" class=button>
	<INPUT type=button value="Cancel" class=button onclick="javascript:window.history.go(-1);">
</P>
</FORM>
<%
	rec.Close
	con.Close
	set rec = nothing
	set con = nothing
%>
		
		</TD>
	</TR>
</TABLE>

<P>&nbsp;</P>

</BODY>
</HTML>
<SCRIPT LANGUAGE=vbscript RUNAT=Server>
Function GetTypeName(intType)
	Select Case intType
	Case 3		GetTypeName = "Long Integer"
	Case 7		GetTypeName = "Date/Time"
	Case 11		GetTypeName = "Boolean"
	Case 6		GetTypeName = "Currency"
	Case 128,204	GetTypeName = "Binary"
	Case 17		GetTypeName = "Byte"
	Case 131	GetTypeName = "Decimal"
	Case 5		GetTypeName = "Double"
	Case 2		GetTypeName = "Integer"
	Case 4		GetTypeName = "Single"
	Case 72		GetTypeName = "Replication ID"
	Case 203,201 GetTypeName = "Memo"
	Case 202,200 GetTypeName = "Text"
	Case Else	GetTypeName = intType
	End Select
End Function
</SCRIPT>
