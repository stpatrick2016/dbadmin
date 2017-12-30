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
		
		XMLFilePath_ = Server.MapPath(FilePath)
		xmlDoc_.load(XMLFilePath_)
		if xmlDoc_.parseError.errorCode <> 0 then 
			'check if the file exists before creating a new one
			dim fso, xml
			set fso = Server.CreateObject("Scripting.FileSystemObject")
			xml = "<?xml version=""1.0"" ?><spp><common /><users>"
			if Len(Username) > 0 then xml = xml & "<" & Username & " />"
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
			GetProfileString = Cstr(Node.text)
			
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
			set ParentNode = ParentNode.appendChild(xmlDoc_.createElement(Username_))
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

'***** Private members **************'
	Private XMLFilePath_
	Private Username_
	Private xmlDoc_

	Private Sub Class_Initialize
		XMLFilePath_ = ""
		Username_ = ""
		LastError = ""
		
		set xmlDoc_ = Server.CreateObject("Msxml2.DOMDocument")
		xmlDoc_.async = False
	End Sub

	Private Sub Class_Terminate
		set xmlDoc_ = Nothing
	End Sub
	
	Private Function IsInitialized
		if Len(XMLFilePath_) > 0 and xmlDoc_.parseError.errorCode = 0 then IsInitialized = True else IsInitialized = False
	End Function
	
	Private Function BuildPath(RelativePath)
		dim path
		path = "/spp"
		if Len(Username_) > 0 then path = path & "/users/" & Username_ else path = path & "/common"
		if Len(RelativePath) > 0 then path = path & "/" & RelativePath
		
		BuildPath = path
	End Function
	

End Class

%>