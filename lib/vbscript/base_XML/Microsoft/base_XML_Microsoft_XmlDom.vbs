Option Explicit

Class base_XML_Microsoft_XmlDom
	Private p_objXmlDom

	Private Sub Class_Initialize()
		Set p_objXmlDom = CreateObject("Microsoft.XMLDOM.1.0")
	End Sub


	' Properties


	Public Property Get Async()
		Async = p_objXmlDom.async 
	End Property

	Public Property Let Async( _
		ByVal blnAsync _
		)

		p_objXmlDom.async = blnAsync
	End Property

	Public Property Get Attributes()
		Set Attributes = p_objXmlDom.attributes 
	End Property

	Public Property Get BaseName() 
		BaseName = p_objXmlDom.baseName 
	End Property

	Public Property Get ChildNodes()
		Set ChildNodes = p_objXmlDom.childNodes 
	End Property

	Public Property Get DataType()
		If IsObject(p_objXmlDom.dataType) Then
			Set DataType = p_objXmlDom.dataType
		Else
			DataType = p_objXmlDom.dataType
		End If
	End Property

	Public Property Let DataType( _
		ByVal varDataType _
		)

		p_objXmlDom.dataType = varDataType 
	End Property

	Public Property Set DataType( _
		ByVal varDataType _
		)

		Set p_objXmlDom.dataType = varDataType
	End Property

	Public Property Get Definition()
		Set Definition = p_objXmlDom.definition 
	End Property

	Public Property Get Doctype()
		Set Doctype = p_objXmlDom.doctype 
	End Property

	Public Property Get DocumentElement()
		Set DocumentElement = p_objXmlDom.documentElement 
	End Property

	Public Property Set DocumentElement( _
		ByVal objDocumentElement _
		)

		Set p_objXmlDom.documentElement = objDocumentElement
	End Property

	Public Property Get FirstChild()
		Set FirstChild = p_objXmlDom.firstChild 
	End Property
 
	Public Property Get Implementation()
		Set Implementation = p_objXmlDom.implementation 
	End Property

	Public Property Get LastChild()
		Set LastChild = p_objXmlDom.lastChild 
	End Property

	Public Property Get Namespaces()
		Set Namespaces = p_objXmlDom.namespaces 
	End Property

	Public Property Get NamespaceUri()
		NamespaceUri = p_objXmlDom.namespaceURI 
	End Property
 
	Public Property Get NodeName()
		NodeName = p_objXmlDom.nodeName 
	End Property

	Public Property Get NodeType()
		Set NodeType = p_objXmlDom.nodeType 
	End Property

	Public Property Get NodeTypedValue()
		If IsObject(p_objXmlDom.nodeTypedValue) Then
			Set NodeTypedValue = p_objXmlDom.nodeTypedValue
		Else
			NodeTypedValue = p_objXmlDom.nodeTypedValue
		End If 
	End Property

	Public Property Let NodeTypedValue( _
		ByVal varNodeTypedValue _
		)

		p_objXmlDom.nodeTypedValue = varNodeTypedValue
	End Property

	Public Property Set NodeTypedValue( _
		ByVal varNodeTypedValue _
		)

		Set p_objXmlDom.nodeTypedValue = varNodeTypedValue
	End Property

	Public Property Get NodeTypeString()
		NodeTypeString = p_objXmlDom.nodeTypeString 
	End Property

	Public Property Get NodeValue()
		If IsObject(p_objXmlDom.nodeValue) Then
			Set NodeValue = p_objXmlDom.nodeValue
		Else
			NodeValue = p_objXmlDom.nodeValue
		End If 
	End Property

	Public Property Let NodeValue( _
		ByVal varNodeValue _
		)

		p_objXmlDom.nodeValue = varNodeValue
	End Property

	Public Property Set NodeValue( _
		ByVal varNodeValue _
		)

		Set p_objXmlDom.nodeValue = varNodeValue
	End Property

	Public Property Get OnDataAvailable()
		If IsObject(p_objXmlDom.ondataavailable) Then
			Set OnDataAvailable = p_objXmlDom.ondataavailable
		Else
			OnDataAvailable = p_objXmlDom.ondataavailable
		End If 
	End Property

	Public Property Let OnDataAvailable( _
		ByVal varOnDataAvailable _
		)

		p_objXmlDom.ondataavailable = varOnDataAvailable
	End Property

	Public Property Set OnDataAvailable( _
		ByVal varOnDataAvailable _
		)

		Set p_objXmlDom.ondataavailable = varOnDataAvailable
	End Property

	Public Property Get OnReadyStateChange()
		If IsObject(p_objXmlDom.onreadystatechange) Then
			Set OnReadyStateChange = p_objXmlDom.onreadystatechange
		Else
			OnReadyStateChange = p_objXmlDom.onreadystatechange
		End If 
	End Property

	Public Property Let OnReadyStateChange( _
		ByVal varOnReadyStateChange _
		)

		p_objXmlDom.onreadystatechange = varOnReadyStateChange
	End Property

	Public Property Set OnReadyStateChange( _
		ByVal varOnReadyStateChange _
		)

		Set p_objXmlDom.onreadystatechange = varOnReadyStateChange
	End Property

	Public Property Get OnTransformNode()
		If IsObject(p_objXmlDom.ontransformnode) Then
			Set OnTransformNode = p_objXmlDom.ontransformnode
		Else
			OnTransformNode = p_objXmlDom.ontransformnode
		End If 
	End Property

	Public Property Let OnTransformNode( _
		ByVal varOnTransformNode _
		)

		p_objXmlDom.ontransformnode = varOnTransformNode
	End Property

	Public Property Set OnTransformNode( _
		ByVal varOnTransformNode _
		)

		Set p_objXmlDom.ontransformnode = varOnTransformNode
	End Property

	Public Property Get OwnerDocument()
		Set OwnerDocument = p_objXmlDom.ownerDocument 
	End Property
 
	Public Property Get ParentNode()
		Set ParentNode = p_objXmlDom.parentNode 
	End Property

	Public Property Get Parsed()
		Parsed = p_objXmlDom.parsed 
	End Property

	Public Property Get ParseError()
		Set ParseError = p_objXmlDom.parseError 
	End Property

	Public Property Get Prefix()
		Prefix = p_objXmlDom.prefix 
	End Property

	Public Property Get PreserveWhiteSpace()
		PreserveWhiteSpace = p_objXmlDom.preserveWhiteSpace 
	End Property

	Public Property Let PreserveWhiteSpace( _
		ByVal blnPreserveWhiteSpace _
		)

		p_objXmlDom.preserveWhiteSpace = blnPreserveWhiteSpace
	End Property

	Public Property Get PreviousSibling()
		Set PreviousSibling = p_objXmlDom.previousSibling 
	End Property

	Public Property Get ReadyState()
		ReadyState = p_objXmlDom.readyState 
	End Property

	Public Property Get ResolveExternals()
		ResolveExternals = p_objXmlDom.resolveExternals 
	End Property

	Public Property Let ResolveExternals( _
		ByVal blnResolveExternals _
		)

		p_objXmlDom.resolveExternals = blnResolveExternals
	End Property
 
	Public Property Get Schemas()
		If IsObject(p_objXmlDom.schemas) Then
			Set Schemas = p_objXmlDom.schemas
		Else
			Schemas = p_objXmlDom.schemas
		End If 
	End Property

	Public Property Let Schemas( _
		ByVal varSchemas _
		)

		p_objXmlDom.schemas = varSchemas
	End Property

	Public Property Set Schemas( _
		ByVal varSchemas _
		)

		Set p_objXmlDom.schemas = varSchemas
	End Property

	Public Property Get Specified()
		Specified = p_objXmlDom.specified 
	End Property

	Public Property Get Text()
		Text = p_objXmlDom.text 
	End Property

	Public Property Let Text( _
		ByVal strText _
		)

		p_objXmlDom.text = strText
	End Property

	Public Property Get Url()
		Url = p_objXmlDom.url 
	End Property

	Public Property Get ValidateOnParse()
		ValidateOnParse = p_objXmlDom.validateOnParse 
	End Property

	Public Property Let ValidateOnParse( _
		ByVal blnValidateOnParse _
		)

		p_objXmlDom.validateOnParse = blnValidateOnParse
	End Property

	Public Property Get Xml()
		Xml = p_objXmlDom.xml 
	End Property


	' Methods


	Public Sub Abort()
		p_objXmlDom.abort
	End Sub

	Public Function AppendChild( _
		ByVal objNewChild _
		)

		Set AppendChild = p_objXmlDom.appendChild(objNewChild)
	End Function

	Public Function CloneNode( _
		ByVal blnDeep _
		)

		Set CloneNode = p_objXmlDom.cloneNode(blnDeep)
	End Function

	Public Function CreateAttribute( _
		ByVal strName _
		)

		Set CreateAttribute = p_objXmlDom.createAttribute(strName)
	End Function
 
	Public Function CreateCdataSection( _
		ByVal strData _
		)

		Set CreateCdataSection = p_objXmlDom.createCDATASection(strData)
	End Function
 
	Public Function CreateComment( _
		ByVal strData _
		)

		Set CreateComment = p_objXmlDom.createComment(strData)
	End Function
 
	Public Function CreateDocumentFragment() 
		Set CreateDocumentFragment = p_objXmlDom.createDocumentFragment()
	End Function

	Public Function CreateElement( _
		ByVal strTagName _
		)

		Set CreateElement = p_objXmlDom.createElement(strTagName)
	End Function
 
	Public Function CreateEntityReference( _
		ByVal strName _
		)

		Set CreateEntityReference = p_objXmlDom.createEntityReference(strName)
	End Function

	Public Function CreateNode( _
		ByVal varType, _
		ByVal strName, _
		ByVal strNamespaceUri _
		)

		Set CreateNode = p_objXmlDom.createNode(varType, strName, strNamespaceUri)
	End Function
 
	Public Function CreateProcessingInstruction( _
		ByVal strTarget, _
		ByVal strData _
		)
 
		Set CreateProcessingInstruction = p_objXmlDom.createProcessingInstruction(strTarget, strData) 
	End Function

	Public Function CreateTextNode( _
		ByVal strData _
		)

		Set CreateTextNode = p_objXmlDom.createTextNode(strData)
	End Function

	Public Function GetElementsByTagName( _
		ByVal strTagName _
		)

		Set GetElementsByTagName = p_objXmlDom.getElementsByTagName(strTagName)
	End Function
 
	Public Function GetProperty( _
		ByVal strName _
		)

		GetProperty = p_objXmlDom.getProperty(strName)
	End Function

	Public Function HasChildNodes()
		HasChildNodes = p_objXmlDom.hasChildNodes()
	End Function

	Public Function InsertBefore( _
		ByVal objNewChild, _
		ByVal varRefChild _
		)

		Set InsertBefore = p_objXmlDom.insertBefore(objNewChild, varRefChild)
	End Function
 
	Public Function Load( _
		ByVal varXmlSource _
		)

		Load = p_objXmlDom.load(varXmlSource)
	End Function

	Public Function LoadXml( _
		ByVal strXml _
		)
 
		LoadXml = p_objXmlDom.loadXML(strXml) 
	End Function

	Public Function NodeFromId( _
		ByVal strIdString _
		)

		Set NodeFromId = p_objXmlDom.nodeFromID(strIdString)
	End Function

	Public Function RemoveChild( _
		ByVal objChildNode _
		)

		Set RemoveChild = p_objXmlDom.removeChild(objChildNode)
	End Function

	Public Function ReplaceChild( _
		ByVal objNewChild, _
		ByVal objOldChild _
		)
 
		Set ReplaceChild = p_objXmlDom.replaceChild(objNewChild, objOldChild)
	End Function

	Public Sub Save( _
		ByVal varDestination _
		)
 
		p_objXmlDom.save varDestination
	End Sub

	Public Function SelectNodes( _
		ByVal strQueryString _
		)

		Set SelectNodes = p_objXmlDom.selectNodes(strQueryString)
	End Function

	Public Function SelectSingleNode( _
		ByVal strQueryString _
		)

		Set SelectSingleNode = p_objXmlDom.selectSingleNode(strQueryString)
	End Function
 
	Public Sub SetProperty( _
		ByVal strName, _
		ByVal varValue _
		)

		p_objXmlDom.setProperty strName, varValue
	End Sub

	Public Function TransformNode( _
		ByVal objStylesheet _
		)

		TransformNode = p_objXmlDom.transformNode(objStylesheet)
	End Function
 
	Public Sub TransformNodeToObject( _
		ByVal objStylesheet, _
		ByVal varOutputObject _
		)

		p_objXmlDom.transformNodeToObject objStylesheet, varOutputObject
	End Sub
 
	Public Function Validate()
		Set Validate = p_objXmlDom.validate()
	End Function

	Private Sub Class_Terminate()
		Set p_objXmlDom = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_XML_Microsoft_XmlDom.vbs" Then

End If