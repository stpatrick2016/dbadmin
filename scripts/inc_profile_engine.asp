<%

'******* Constants ******************'

Class StpPrivateProfile

	Public LastError

'***** Public methods ***************'

	'########################################
	'# Loads an XML file
	Public Function Load(FilePath, Username)
		Load = False
		if Len(FilePath) = 0 then Exit Function
		
		On Error Resume Next
		if Mid(FilePath, 2, 1) = ":" Then XMLFilePath_ = FilePath Else XMLFilePath_ = Server.MapPath(FilePath)
		xmlDoc_.load(XMLFilePath_)
		If Err Then 
			LastError = Err.Description
			Exit Function
		End If
		if xmlDoc_.parseError.errorCode <> 0 then 
			'check if the file exists before creating a new one
			dim fso, xml
			set fso = Server.CreateObject("Scripting.FileSystemObject")
			xml = "<?xml version=""1.0"" ?><spp><common /><users>"
			if Len(Username) > 0 then xml = xml & "<user name=""" & Username & """ />"
			xml = xml & "</users></spp>"
			if not fso.FileExists(XMLFilePath_) then xmlDoc_.loadXML xml
			set fso = Nothing
			
			if xmlDoc_.parseError.errorCode <> 0 then 
				LastError = xmlDoc_.parseError.reason
				XMLFilePath_ = ""
				Exit Function
			end if
		end if
		Username_ = Username
		Load = True
	End Function
	
	'########################################
	'# Saves changes into the file
	Public Function Save
		Save = False
		if Len(XMLFilePath_) = 0 then Exit Function
		
		On Error Resume Next
		xmlDoc_.save(XMLFilePath_)
		If Err then Save = False Else Save = True
	End Function
	
	'########################################
	'# returns a variable as string
	Public Function GetProfileString(Section, Attribute, DefaultValue)
		GetProfileString = DefaultValue
		if not IsInitialized then Exit Function

		dim Node
		set Node = xmlDoc_.selectSingleNode(BuildPath(Section))
		if TypeName(Node) <> "Nothing" then
			if Len(Attribute) > 0 then set Node = Node.attributes.getNamedItem(Attribute)
			if TypeName(Node) <> "Nothing" Then GetProfileString = Cstr(Node.text)
			
			set Node = Nothing
		end if		
	End Function
	
	'########################################
	'# returns a variable as number
	Public Function GetProfileNumber(Section, Attribute, DefaultValue)
		GetProfileNumber = DefaultValue
		if not IsInitialized then Exit Function

		dim Node
		set Node = xmlDoc_.selectSingleNode(BuildPath(Section))
		if TypeName(Node) <> "Nothing" then
			if Len(Attribute) > 0 then set Node = Node.attributes.getNamedItem(Attribute)
			if TypeName(Node) <> "Nothing" then
				if IsNumeric(Node.text) then GetProfileNumber = CLng(Node.text)
			end if
			
			set Node = Nothing
		end if		
	End Function
	
	'########################################
	'# returns all child items as an array
	Public Function GetProfileArray(Section, Attribute)
		dim ret
		ret = Array()
		GetProfileArray = ret
		if not IsInitialized then Exit Function

		dim Nodes, i, Node
		set Nodes = xmlDoc_.selectSingleNode(BuildPath(Section))
		if TypeName(Nodes) <> "Nothing" then 
			set Nodes = Nodes.childNodes
			for i=0 to Nodes.length
				if Len(Attribute) > 0 then set Node = Nodes(i).attributes.getNamedItem(Attribute) else set Node = Nodes(i)
				if TypeName(Node) <> "Nothing" then
					if Len(Attribute) > 0 then set Node = Node.attributes.getNamedItem(Attribute)
					redim preserve ret(i)
					ret(i) = CStr(Node.text)
					
					set Node = Nothing
				end if		
			next
		end if
		GetProfileArray = ret
	End Function
	
	'########################################
	'# sets a new value 
	Public Sub SetValue(Section, Attribute, Value)
		if not IsInitialized then Exit Sub

		dim Node, Sections, i, ParentNode, Attr, ParentPath
		Sections = Split(Section, "/")
		set ParentNode = xmlDoc_.selectSingleNode(BuildPath(""))
		
		'build user's node
		if TypeName(ParentNode) = "Nothing" then
			dim tempPath : tempPath = BuildPath("")
			tempPath = Left(tempPath, InStrRev(tempPath, "/") - 1)
			set ParentNode = xmlDoc_.selectSingleNode(tempPath)
			set ParentNode = ParentNode.appendChild(xmlDoc_.createElement("user"))
			set Attr = xmlDoc_.createAttribute("name")
			Attr.value = Username_
			call ParentNode.attributes.setNamedItem(Attr)
			set Attr = Nothing
		end if
		
		ParentPath = BuildPath("")
		for i=0 to UBound(Sections)
			ParentPath = ParentPath & "/" & Sections(i)
			set Node = xmlDoc_.selectSingleNode(ParentPath)
			if TypeName(Node) = "Nothing" then 
				set Node = ParentNode.appendChild(xmlDoc_.createElement(Sections(i)))
			end if
			set ParentNode = Node
		next
		
		'now we have all path created and ready
		if IsArray(Value) then
			do while Node.childNodes.length > 0 
				Node.removeChild Node.childNodes(0)
			loop
			set ParentNode = Node
			for i=0 to ubound(Value)
				set Node = ParentNode.appendChild(xmlDoc_.createElement("item"))
				if Len(Attribute) > 0 then
					set Attr = xmlDoc_.createAttribute(Attribute)
					Attr.value = Value(i)
					Node.attributes.setNamedItem(Attr)
					set Attr = Nothing
				else
					Node.appendChild xmlDoc_.createCDATASection(Value(i))
				end if
			next
		else
			if Len(Attribute) > 0 then 
				set ParentNode = Node
				set Node = xmlDoc_.createAttribute(Attribute)
				Node.value = Value
				ParentNode.attributes.setNamedItem(Node)
			else
				if Node.childNodes.length > 0 then Node.removeChild Node.childNodes(0)
				Node.appendChild xmlDoc_.createCDATASection(Value)
			end if
		end if
	End Sub

	'########################################
	'# Removes given node
	Public Function RemoveNode(XPath)
		if not IsInitialized then Exit Function

		dim Node, Parent
		XPath = BuildPath(XPath)
		Set Node = xmlDoc_.selectSingleNode(XPath)
		If not Node is Nothing Then
			Set Parent = Node.parentNode
			call Parent.removeChild(Node)
		End If
		
		Set Parent = Nothing
		Set Node = Nothing
		RemoveNode = True
	End Function
	
	'########################################
	'# Returns either cookie or Session variable, regarding of settings
	Public Function GetCookie(key)
		dim bUseCookies, strTemp, strPassword
		
		strTemp = Username_
		Username_ = ""
		if Me.GetProfileNumber("settings", "use_cookies", 0) <> 0 then bUseCookies = True else bUseCookies = False
		Username_ = strTemp
		
		if bUseCookies then
			strPassword = Request.Cookies("DBAdmin")("password")
			
		else
			GetCookie = CStr(Session(key))
		end if
	End Function 
	
	'########################################
	'# Returns True is a given component is available
	Public Function ComponentAvailable(Component)
		Dim ProgID, Obj
		Select Case ucase(Component)
			Case "ADOX"		ProgID = "ADOX.Catalog"
			Case "ADO"		ProgID = "ADODB.Connection"
			Case "XML3"		ProgID = "MSXML.DOMDocument"
			Case "XML4"		ProgID = "MSXML.DOMDocument.4"
			Case Else		ProgID = Component
		End Select
		Set Obj = Server.CreateObject(ProgID)
		If IsEmpty(Obj) or Obj Is Nothing Then ComponentAvailable = False Else ComponentAvailable = True
	End Function

'***** Private members **************'
	Private XMLFilePath_
	Private Username_
	Private xmlDoc_

	Private Sub Class_Initialize
		XMLFilePath_ = ""
		Username_ = ""
		LastError = ""
		
		On Error Resume Next
		'lets see if user has set it to his own progID
		If IsEmpty(DBA_cfgMSXMLProgID) Or Len(DBA_cfgMSXMLProgID) = 0 Then
			'first try to create MSXML4
			set xmlDoc_ = Server.CreateObject("Msxml2.DOMDocument.4")
			'if not available then try to create version 3
			if xmlDoc_ is Nothing then set xmlDoc_ = Server.CreateObject("Msxml2.DOMDocument")
			'if not available again - well generic form, last chance
			if xmlDoc_ is Nothing then set xmlDoc_ = Server.CreateObject("Microsoft.XMLDOM")
		Else
			Set xmlDoc_ = Server.CreateObject(DBA_cfgMSXMLProgID)
		End If
		if not xmlDoc_ is Nothing then xmlDoc_.async = False
	End Sub

	Private Sub Class_Terminate
		set xmlDoc_ = Nothing
	End Sub
	
	Private Function IsInitialized
		On Error Resume Next
		If TypeName(xmlDoc_) = "Nothing" Then IsInitialized = False
		if Len(XMLFilePath_) > 0 and xmlDoc_.parseError.errorCode = 0 then IsInitialized = True else IsInitialized = False
	End Function
	
	Private Function BuildPath(RelativePath)
		dim path
		path = "/spp"
		if Len(Username_) > 0 then path = path & "/users/user[@name=""" & Username_ & """]" else path = path & "/common"
		if Len(RelativePath) > 0 then path = path & "/" & RelativePath
		
		BuildPath = path
	End Function
	

End Class

%>