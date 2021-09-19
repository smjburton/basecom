Option Explicit

Class base_XML_MSXML_DomDocument
	Private p_objXmlDocument

	Private Sub Class_Initialize()
		Set p_objXmlDocument = CreateObject("MSXML.DOMDocument")
	End Sub


	' Properties


	Public Property Get Async()
		Async = p_objXmlDocument.async
	End Property

	Public Property Let Async( _
		ByVal blnAsync _
		)

		p_objXmlDocument.async = blnAsync
	End Property

	Public Property Get Attributes() 
		Set Attributes = p_objXmlDocument.attributes
	End Property

	Public Property Get BaseName() 
		BaseName = p_objXmlDocument.baseName
	End Property

	Public Property Get ChildNodes()
		Set ChildNodes = p_objXmlDocument.childNodes
	End Property

	Public Property Get DataType()
		If IsObject(p_objXmlDocument.dataType) Then
			Set DataType = p_objXmlDocument.dataType
		Else
			DataType = p_objXmlDocument.dataType
		End If
	End Property

	Public Property Let DataType( _
		ByVal varDataType _
		)

		p_objXmlDocument.dataType = varDataType
	End Property

	Public Property Set DataType( _
		ByVal varDataType _
		)

		Set p_objXmlDocument.dataType = varDataType
	End Property

	Public Property Get Definition()
		Set Definition = p_objXmlDocument.definition
	End Property

	Public Property Get Doctype()
		Set Doctype = p_objXmlDocument.doctype
	End Property

	Public Property Get DocumentElement()
		Set DocumentElement = p_objXmlDocument.documentElement
	End Property

	Public Property Set DocumentElement( _
		ByVal objDocumentElement _
		)

		Set p_objXmlDocument.documentElement = objDocumentElement
	End Property

	Public Property Get FirstChild()
		Set FirstChild = p_objXmlDocument.firstChild
	End Property

	Public Property Get Implementation()
		Set Implementation = p_objXmlDocument.implementation
	End Property

	Public Property Get LastChild() 
		Set LastChild = p_objXmlDocument.lastChild
	End Property

	Public Property Get Namespaces()
		Set Namespaces = p_objXmlDocument.namespaces
	End Property

	Public Property Get NamespaceUri()
		NamespaceUri = p_objXmlDocument.namespaceURI
	End Property

	Public Property Get NodeName()
		NodeName = p_objXmlDocument.nodeName
	End Property

	Public Property Get NodeType()
		Set NodeType = p_objXmlDocument.nodeType
	End Property

	Public Property Get NodeTypedValue()
		If IsObject(p_objXmlDocument.nodeTypedValue) Then
			Set NodeTypedValue = p_objXmlDocument.nodeTypedValue
		Else
			NodeTypedValue = p_objXmlDocument.nodeTypedValue
		End If
	End Property

	Public Property Let NodeTypedValue( _
		ByVal varNodeTypedValue _
		)

		p_objXmlDocument.nodeTypedValue = varNodeTypedValue
	End Property

	Public Property Set NodeTypedValue( _
		ByVal varNodeTypedValue _
		)

		Set p_objXmlDocument.nodeTypedValue = varNodeTypedValue
	End Property

	Public Property Get NodeTypeString()
		NodeTypeString = p_objXmlDocument.nodeTypeString
	End Property

	Public Property Get NodeValue()
		If IsObject(p_objXmlDocument.nodeValue) Then
			Set NodeValue = p_objXmlDocument.nodeValue
		Else
			NodeValue = p_objXmlDocument.nodeValue
		End If
	End Property

	Public Property Let NodeValue( _
		ByVal varNodeValue _
		)

		p_objXmlDocument.nodeValue = varNodeValue
	End Property

	Public Property Set NodeValue( _
		ByVal varNodeValue _
		)

		Set p_objXmlDocument.nodeValue = varNodeValue
	End Property

	Public Property Get OnDataAvailable()
		If IsObject(p_objXmlDocument.ondataavailable) Then
			Set OnDataAvailable = p_objXmlDocument.ondataavailable
		Else
			OnDataAvailable = p_objXmlDocument.ondataavailable
		End If
	End Property

	Public Property Let OnDataAvailable( _
		ByVal varOnDataAvailable _
		)

		p_objXmlDocument.ondataavailable = varOnDataAvailable
	End Property

	Public Property Set OnDataAvailable( _
		ByVal varOnDataAvailable _
		)

		Set p_objXmlDocument.ondataavailable = varOnDataAvailable
	End Property

	Public Property Get OnReadyStateChange()
		If IsObject(p_objXmlDocument.onreadystatechange) Then
			Set OnReadyStateChange = p_objXmlDocument.onreadystatechange
		Else
			OnReadyStateChange = p_objXmlDocument.onreadystatechange
		End If
	End Property

	Public Property Let OnReadyStateChange( _
		ByVal varOnReadyStateChange _
		)
 
		p_objXmlDocument.onreadystatechange = varOnReadyStateChange
	End Property

	Public Property Set OnReadyStateChange( _
		ByVal varOnReadyStateChange _
		)

		Set p_objXmlDocument.onreadystatechange = varOnReadyStateChange
	End Property

	Public Property Get OnTransformNode()
		If IsObject(p_objXmlDocument.ontransformnode) Then
			Set OnTransformNode = p_objXmlDocument.ontransformnode
		Else
			OnTransformNode = p_objXmlDocument.ontransformnode
		End If
	End Property

	Public Property Let OnTransformNode( _
		ByVal varOnTransformNode _
		)

		p_objXmlDocument.ontransformnode = varOnTransformNode
	End Property

	Public Property Set OnTransformNode( _
		ByVal varOnTransformNode _
		)

		Set p_objXmlDocument.ontransformnode = varOnTransformNode
	End Property

	Public Property Get OwnerDocument()
		Set OwnerDocument = p_objXmlDocument.ownerDocument
	End Property

	Public Property Get ParentNode()
		Set ParentNode = p_objXmlDocument.parentNode
	End Property

	Public Property Get Parsed()
		Parsed = p_objXmlDocument.parsed
	End Property

	Public Property Get ParseError()
		Set ParseError = p_objXmlDocument.parseError
	End Property

	Public Property Get Prefix()   
		Prefix = p_objXmlDocument.prefix
	End Property

	Public Property Get PreserveWhiteSpace()
		PreserveWhiteSpace = p_objXmlDocument.preserveWhiteSpace
	End Property

	Public Property Let PreserveWhiteSpace( _
		ByVal blnPreserveWhiteSpace _
		)

		p_objXmlDocument.preserveWhiteSpace = blnPreserveWhiteSpace
	End Property

	Public Property Get PreviousSibling()
		Set PreviousSibling = p_objXmlDocument.previousSibling
	End Property

	Public Property Get ReadyState()
		ReadyState = p_objXmlDocument.readyState
	End Property

	Public Property Get ResolveExternals()
		ResolveExternals = p_objXmlDocument.resolveExternals
	End Property

	Public Property Let ResolveExternals( _
		ByVal blnResolveExternals _
		)

		p_objXmlDocument.resolveExternals = blnResolveExternals
	End Property

	Public Property Get Schemas()
		If IsObject(p_objXmlDocument.schemas) Then
			Set Schemas = p_objXmlDocument.schemas
		Else
			Schemas = p_objXmlDocument.schemas
		End If
	End Property

	Public Property Let Schemas( _
		ByVal varSchemas _
		)

		p_objXmlDocument.schemas = varSchemas
	End Property

	Public Property Set Schemas( _
		ByVal varSchemas _
		)

		p_objXmlDocument.schemas = varSchemas
	End Property

	Public Property Get Specified() 
		Specified = p_objXmlDocument.specified
	End Property

	Public Property Get Text()
		Text = p_objXmlDocument.text
	End Property

	Public Property Let Text( _
		ByVal strText _
		)

		p_objXmlDocument.text = strText
	End Property

	Public Property Get Url()
		Url = p_objXmlDocument.url
	End Property

	Public Property Get ValidateOnParse()
		ValidateOnParse = p_objXmlDocument.validateOnParse
	End Property

	Public Property Let ValidateOnParse( _
		ByVal blnValidateOnParse _
		)

		p_objXmlDocument.validateOnParse = blnValidateOnParse
	End Property

	Public Property Get Xml()
		Xml = p_objXmlDocument.xml
	End Property


	' Methods


	Public Sub Abort()
		p_objXmlDocument.abort
	End Sub

	Public Function AppendChild( _
		ByVal objNewChild _
		)

		Set AppendChild = p_objXmlDocument.appendChild(objNewChild)
	End Function

	Public Function CloneNode( _
		ByVal blnDeep _
		)

		Set CloneNode = p_objXmlDocument.cloneNode(blnDeep)
	End Function

	Public Function CreateAttribute( _
		ByVal strName _
		)

		Set CreateAttribute = p_objXmlDocument.createAttribute(strName)
	End Function

	Public Function CreateCdataSection( _
		ByVal strData _
		)
 
		Set CreateCdataSection = p_objXmlDocument.createCDATASection(strData)
	End Function

	Public Function CreateComment( _
		ByVal strData _
		)

		Set CreateComment = p_objXmlDocument.createComment(strData)
	End Function

	Public Function CreateDocumentFragment()
		Set CreateDocumentFragment = p_objXmlDocument.createDocumentFragment()
	End Function

	Public Function CreateElement( _
		ByVal strTagName _
		)

		Set CreateElement = p_objXmlDocument.createElement(strTagName)
	End Function

	Public Function CreateEntityReference( _
		ByVal strName _
		)

		Set CreateEntityReference = p_objXmlDocument.createEntityReference(strName)
	End Function

	Public Function CreateNode( _
		ByVal varType, _
		ByVal strName, _
		ByVal strNamespaceUri _
		)

		Set CreateNode = p_objXmlDocument.createNode(varType, strName, strNamespaceUri)
	End Function

	Public Function CreateProcessingInstruction( _
		ByVal strTarget, _
		ByVal strData _
		)

		Set CreateProcessingInstruction = p_objXmlDocument.createProcessingInstruction(strTarget, strData)
	End Function

	Public Function CreateTextNode( _
		ByVal strData _
		)

		Set CreateTextNode = p_objXmlDocument.createTextNode(strData)
	End Function

	Public Function GetElementsByTagName( _
		ByVal strTagName _
		)

		Set GetElementsByTagName = p_objXmlDocument.getElementsByTagName(strTagName)
	End Function

	Public Function GetProperty( _
		ByVal strName _
		)

		GetProperty = p_objXmlDocument.getProperty(strName)
	End Function

	Public Function HasChildNodes()
		HasChildNodes = p_objXmlDocument.hasChildNodes()
	End Function

	Public Function InsertBefore( _
		ByVal objNewChild, _
		ByVal varRefChild _
		)

		Set InsertBefore = p_objXmlDocument.insertBefore(objNewChild, varRefChild)
	End Function

	Public Function Load( _
		ByVal varXmlSource _
		)

		Load = p_objXmlDocument.load(varXmlSource)
	End Function

	Public Function LoadXml( _
		ByVal strXml _
		)

		LoadXml = p_objXmlDocument.loadXML(strXml)
	End Function

	Public Function NodeFromId( _
		ByVal strIdString _
		)

		Set NodeFromId = p_objXmlDocument.nodeFromID(strIdString)
	End Function

	Public Function RemoveChild( _
		ByVal objChildNode _
		)

		Set RemoveChild = p_objXmlDocument.removeChild(objChildNode)
	End Function

	Public Function ReplaceChild( _
		ByVal objNewChild, _
		ByVal objOldChild _
		)

		Set ReplaceChild = p_objXmlDocument.replaceChild(objNewChild, objOldChild)
	End Function

	Public Sub Save( _
		ByVal varDestination _
		)
 
		p_objXmlDocument.save varDestination
	End Sub

	Public Function SelectNodes( _
		ByVal strQueryString _
		)

		Set SelectNodes = p_objXmlDocument.selectNodes(strQueryString)
	End Function

	Public Function SelectSingleNode( _
		ByVal strQueryString _
		)

		Set SelectSingleNode = p_objXmlDocument.selectSingleNode(strQueryString)
	End Function

	Public Sub SetProperty( _
		ByVal strName, _
		ByVal varValue _
		)

		p_objXmlDocument.setProperty strName, varValue
	End Sub
 
	Public Function TransformNode( _
		ByVal objStylesheet _
		)

		TransformNode = p_objXmlDocument.transformNode(objStylesheet)
	End Function

	Public Sub TransformNodeToObject( _
		ByVal objStylesheet, _
		ByVal varOutputObject _
		)

		p_objXmlDocument.transformNodeToObject objStylesheet, varOutputObject
	End Sub

	Public Function Validate()
		Set Validate = p_objXmlDocument.validate()
	End Function

	Private Sub Class_Terminate()
		Set p_objXmlDocument = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_XML_MSXML_DomDocument.vbs" Then

End If