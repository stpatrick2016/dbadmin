<%
'######################################################################################
'# CMultiLangEngine - a class for multi language support for ASP applications
'# Copyright © 2002-2003 by Philip Patrick. All rights reserved
'# E-mail: stpatrick@mail.com
'# Web-site: http://www.stpworks.com
'#
'# This code cannot be used and/or redistributed in any form without permission of its author(s) 
'#
'# DISCLAIMER:
'#	This software is provided 'AS IS' and any express or implied				#
'#  warranties, including, but not limited to, the implied warranties of		#
'#  merchantability and fitness for a particular purpose, are disclaimed.		#
'#  In no event shall the authors be liable for any direct, indirect,			#
'#  incidental, special, exemplary, or consequential damages (including, but	#
'#  not limited to, procurement of substitute goods or services; loss of use,	#
'#  data, or profits; or business interruption) however caused and on any		#
'#  theory of liability, whether in contract, strict liability, or tort			#
'#  (including negligence or otherwise) arising in any way out of the use of	#
'#  this software, even if advised of the possibility of such damage.			#

Class CMultiLangEngine

	Public Default Property Get Header(key)
		if m_dicHeader.Exists(key) Then Header = m_dicHeader(key)
	End Property

	'###############################################
	'# Loads a text file with specified FilePath. If VarPrefix specified appends a prefix to
	'# each variable name. If HeaderOnly is True then no strings loaded except the file header
	Public Sub Load(FilePath, VarPrefix, HeaderOnly)
		dim fso, ts, line, arr, script
		
		On Error Resume Next
		If Len(FilePath) = 0 Then Exit Sub
		if Mid(FilePath, 2, 1) <> ":" Then FilePath = Server.MapPath(FilePath)
		m_dicHeader.RemoveAll
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		set ts = fso.OpenTextFile(FilePath, 1, False)	'ForReading
		
		If Err = 0 Then
			Do While not ts.AtEndOfStream
				line = Trim(ts.ReadLine())
				If Len(line) > 1 Then
					If Left(line, 2) = "#@" Then							'header's string
						line = LTrim(Mid(line, 3))
						arr = Split(line, "=", 2)
						call ParseHeaderVariable(arr)
					ElseIf Left(line, 1) <> "#" and not HeaderOnly Then		'not a comment and strings should be loaded
						arr = Split(line, "=", 2)
						If ubound(arr) > 0 Then
							arr(0) = RTrim(arr(0))
							arr(1) = Replace(LTrim(arr(1)), """", """""")
							script = VarPrefix & arr(0) & " = """ & arr(1) & """"
							On Error Resume Next
							call Execute(script)
							On Error Goto 0
						End If
					End If
				End If
			Loop
		End If
		On Error Goto 0
		
		set ts = Nothing
		set fso = Nothing
	End Sub
	
'****************** Private *************************
	
	Private m_dicHeader
	
	Private Sub Class_Initialize
		set m_dicHeader = Server.CreateObject("Scripting.Dictionary")
	End Sub
	
	Private Sub Class_Terminate
		set m_dicHeader = Nothing
	End Sub
	
	Private Sub ParseHeaderVariable(ByRef arr)
		If ubound(arr) < 1 Then Exit Sub
		arr(0) = Trim(arr(0))
		arr(1) = LTrim(arr(1))
		if Len(arr(0)) > 0 Then m_dicHeader(arr(0)) = arr(1)
	End Sub
	
End Class
%>