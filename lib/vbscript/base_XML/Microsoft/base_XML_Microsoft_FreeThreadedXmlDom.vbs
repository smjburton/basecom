Option Explicit

Class base_XML_Microsoft_FreeThreadedXmlDom
	Private p_objFreeThreadedXmlDom

	Private Sub Class_Initialize()
		Set p_objFreeThreadedXmlDom = CreateObject("Microsoft.FreeThreadedXMLDOM.1.0")
	End Sub


	' Properties


	Public Property Get Async()
		Async = p_objFreeThreadedXmlDom.async
	End Property

	Public Property Let Async( _
		ByVal blnAsync _
		)

		p_objFreeThreadedXmlDom.async = blnAsync 
	End Property
 
	Public Property Get Attributes()
		Set Attributes = p_objFreeThreadedXmlDom.attributes 
	End Property

	Public Property Get BaseName()
		BaseName = p_objFreeThreadedXmlDom.baseName
	End Property

	Public Property Get ChildNodes()
		Set ChildNodes = p_objFreeThreadedXmlDom.childNodes 
	End Property

	Public Property Get DataType()
		If IsObject(p_objFreeThreadedXmlDom.dataType) Then
			Set DataType = p_objFreeThreadedXmlDom.dataType
		Else
			DataType = p_objFreeThreadedXmlDom.dataType
		End If
	End Property

	Public Property Let DataType( _
		ByVal varDataType _
		)

		p_objFreeThreadedXmlDom.dataType = varDataType
	End Property

	Public Property Set DataType( _
		ByVal varDataType _
		)

		Set p_objFreeThreadedXmlDom.dataType = varDataType 
	End Property

	Public Property Get Definition()
		Set Definition = p_objFreeThreadedXmlDom.definition 
	End Property

	Public Property Get Doctype() 
		Set Doctype = p_objFreeThreadedXmlDom.doctype 
	End Property

	Public Property Get DocumentElement()
		Set DocumentElement = p_objFreeThreadedXmlDom.documentElement 
	End Property

	Public Property Set DocumentElement( _
		ByVal objDocumentElement _
		)
 
		Set p_objFreeThreadedXmlDom.documentElement = objDocumentElement
	End Property

	Public Property Get FirstChild() 
		Set FirstChild = p_objFreeThreadedXmlDom.firstChild 
	End Property

	Public Property Get Implementation()
		Set Implementation = p_objFreeThreadedXmlDom.implementation 
	End Property

	Public Property Get LastChild()
		Set LastChild = p_objFreeThreadedXmlDom.lastChild 
	End Property

	Public Property Get Namespaces()
		Set Namespaces = p_objFreeThreadedXmlDom.namespaces 
	End Property

	Public Property Get NamespaceURI()
		NamespaceURI = p_objFreeThreadedXmlDom.namespaceURI
	End Property

	Public Property Get NextSibling()
		Set NextSibling = p_objFreeThreadedXmlDom.nextSibling 
	End Property

	Public Property Get NodeName()
		NodeName = p_objFreeThreadedXmlDom.nodeName
	End Property

	Public Property Get NodeType()
		Set NodeType = p_objFreeThreadedXmlDom.nodeType 
	End Property

	Public Property Get NodeTypedValue()
		If IsObject(p_objFreeThreadedXmlDom.nodeTypedValue) Then
			Set NodeTypedValue = p_objFreeThreadedXmlDom.nodeTypedValue
		Else
			NodeTypedValue = p_objFreeThreadedXmlDom.nodeTypedValue
		End If 
	End Property

	Public Property Let NodeTypedValue( _
		ByVal varNodeTypedValue _
		)

		p_objFreeThreadedXmlDom.nodeTypedValue = varNodeTypedValue
	End Property

	Public Property Set NodeTypedValue( _
		ByVal varNodeTypedValue _
		)

		Set p_objFreeThreadedXmlDom.nodeTypedValue = varNodeTypedValue
	End Property

	Public Property Get NodeTypeString()
		NodeTypeString = p_objFreeThreadedXmlDom.nodeTypeString
	End Property

	Public Property Get NodeValue()
		If IsObject(p_objFreeThreadedXmlDom.nodeValue) Then
			Set NodeValue = p_objFreeThreadedXmlDom.nodeValue
		Else
			NodeValue = p_objFreeThreadedXmlDom.nodeValue
		End If
	End Property

	Public Property Let NodeValue( _
		ByVal varNodeValue _
		)

		p_objFreeThreadedXmlDom.nodeValue = varNodeValue
	End Property

	Public Property Set NodeValue( _
		ByVal varNodeValue _
		)

		Set p_objFreeThreadedXmlDom.nodeValue = varNodeValue 
	End Property

	Public Property Get OnDataAvailable()
		If IsObject(p_objFreeThreadedXmlDom.ondataavailable) Then
			Set OnDataAvailable = p_objFreeThreadedXmlDom.ondataavailable
		Else
			OnDataAvailable = p_objFreeThreadedXmlDom.ondataavailable
		End If
	End Property

	Public Property Let OnDataAvailable( _
		ByVal varOnDataAvailable _
		)

		p_objFreeThreadedXmlDom.ondataavailable = varOnDataAvailable 
	End Property

	Public Property Set OnDataAvailable( _
		ByVal varOnDataAvailable _
		)

		Set p_objFreeThreadedXmlDom.ondataavailable = varOnDataAvailable 
	End Property

	Public Property Get OnReadyStateChange()
		If IsObject(p_objFreeThreadedXmlDom.onreadystatechange) Then
			Set OnReadyStateChange = p_objFreeThreadedXmlDom.onreadystatechange
		Else
			OnReadyStateChange = p_objFreeThreadedXmlDom.onreadystatechange
		End If 
	End Property

	Public Property Let OnReadyStateChange( _
		ByVal varOnReadyStateChange _
		)

		p_objFreeThreadedXmlDom.onreadystatechange = varOnReadyStateChange
	End Property

	Public Property Set OnReadyStateChange( _
		ByVal varOnReadyStateChange _
		)

		Set p_objFreeThreadedXmlDom.onreadystatechange = varOnReadyStateChange
	End Property

	Public Property Get OnTransformNode()
		If IsObject(p_objFreeThreadedXmlDom.ontransformnode) Then
			Set OnTransformNode = p_objFreeThreadedXmlDom.ontransformnode
		Else
			OnTransformNode = p_objFreeThreadedXmlDom.ontransformnode
		End If  
	End Property

	Public Property Let OnTransformNode( _
		ByVal varOnTransformNode _
		)

		p_objFreeThreadedXmlDom.ontransformnode = varOnTransformNode
	End Property

	Public Property Set OnTransformNode( _
		ByVal varOnTransformNode _
		)
 
		Set p_objFreeThreadedXmlDom.ontransformnode = varOnTransformNode
	End Property

	Public Property Get OwnerDocument()
		Set OwnerDocument = p_objFreeThreadedXmlDom.ownerDocument 
	End Property

	Public Property Get ParentNode()
		Set ParentNode = p_objFreeThreadedXmlDom.parentNode 
	End Property

	Public Property Get Parsed()
		Parsed = p_objFreeThreadedXmlDom.parsed
	End Property

	Public Property Get ParseError()
		Set ParseError = p_objFreeThreadedXmlDom.parseError 
	End Property

	Public Property Get Prefix()
		Prefix = p_objFreeThreadedXmlDom.prefix
	End Property

	Public Property Get PreserveWhiteSpace()
		PreserveWhiteSpace = p_objFreeThreadedXmlDom.preserveWhiteSpace
	End Property

	Public Property Let PreserveWhiteSpace( _
		ByVal blnPreserveWhiteSpace _
		)

		p_objFreeThreadedXmlDom.preserveWhiteSpace = blnPreserveWhiteSpace
	End Property

	Public Property Get PreviousSibling()
		Set PreviousSibling = p_objFreeThreadedXmlDom.previousSibling 
	End Property

	Public Property Get ReadyState()
		ReadyState = p_objFreeThreadedXmlDom.readyState
	End Property

	Public Property Get ResolveExternals()
		ResolveExternals = p_objFreeThreadedXmlDom.resolveExternals
	End Property

	Public Property Let ResolveExternals( _
		ByVal blnResolveExternals _
		)

		p_objFreeThreadedXmlDom.resolveExternals = blnResolveExternals
	End Property

	Public Property Get Schemas()
		If IsObject(p_objFreeThreadedXmlDom.schemas) Then
			Set Schemas = p_objFreeThreadedXmlDom.schemas
		Else
			Schemas = p_objFreeThreadedXmlDom.schemas
		End If 
	End Property

	Public Property Let Schemas( _
		ByVal varSchemas _
		)

		p_objFreeThreadedXmlDom.schemas = varSchemas
	End Property

	Public Property Set Schemas( _
		ByVal varSchemas _
		)

		Set p_objFreeThreadedXmlDom.schemas = varSchemas 
	End Property

	Public Property Get Specified()
		Specified = p_objFreeThreadedXmlDom.specified
	End Property

	Public Property Get Text()
		Text = p_objFreeThreadedXmlDom.text
	End Property

	Public Property Let Text( _
		ByVal strText _
		)

		p_objFreeThreadedXmlDom.text = strText
	End Property

	Public Property Get Url()
		Url = p_objFreeThreadedXmlDom.url 
	End Property

	Public Property Get ValidateOnParse()
		ValidateOnParse = p_objFreeThreadedXmlDom.validateOnParse
	End Property

	Public Property Let ValidateOnParse( _
		ByVal blnValidateOnParse _
		)
 
		p_objFreeThreadedXmlDom.validateOnParse = blnValidateOnParse
	End Property

	Public Property Get Xml()
		Xml = p_objFreeThreadedXmlDom.xml
	End Property


	' Methods


	Public Sub Abort()
		p_objFreeThreadedXmlDom.abort
	End Sub

	Public Function AppendChild( _
		ByVal objNewChild _
		)

		Set AppendChild = p_objFreeThreadedXmlDom.appendChild(objNewChild)
	End Function

	Public Function CloneNode( _
		ByVal blnDeep _
		)

		Set CloneNode = p_objFreeThreadedXmlDom.cloneNode(blnDeep)
	End Function

	Public Function CreateAttribute( _
		ByVal strName _
		)

		Set CreateAttribute = p_objFreeThreadedXmlDom.createAttribute(strName)
	End Function
 
	Public Function CreateCdataSection( _
		ByVal strData _
		)

		Set CreateCdataSection = p_objFreeThreadedXmlDom.createCDATASection(strData)
	End Function

	Public Function CreateComment( _
		ByVal strData _
		)

		Set CreateComment = p_objFreeThreadedXmlDom.createComment(strData)
	End Function

	Public Function CreateDocumentFragment()
		Set CreateDocumentFragment = p_objFreeThreadedXmlDom.createDocumentFragment()
	End Function
 
	Public Function CreateElement( _
		ByVal strTagName _
		)

		Set CreateElement = p_objFreeThreadedXmlDom.createElement(strTagName)
	End Function
 
	Public Function CreateEntityReference( _
		ByVal strName _
		)

		Set CreateEntityReference = p_objFreeThreadedXmlDom.createEntityReference(strName)
	End Function
 
	Public Function CreateNode( _
		ByVal varType, _
		ByVal strName, _
		ByVal strNamespaceUri _
		)
 
		Set CreateNode = p_objFreeThreadedXmlDom.createNode(varType, strName, strNamespaceUri)
	End Function

	Public Function CreateProcessingInstruction( _
		ByVal strTarget, _
		ByVal strData _
		)
 
		Set CreateProcessingInstruction = p_objFreeThreadedXmlDom.createProcessingInstruction(strTarget, strData)
	End Function

	Public Function CreateTextNode( _
		ByVal strData _
		)

		Set CreateTextNode = p_objFreeThreadedXmlDom.createTextNode(strData)
	End Function

	Public Function GetElementsByTagName( _
		ByVal strTagName _
		)
 
		Set GetElementsByTagName = p_objFreeThreadedXmlDom.getElementsByTagName(strTagName) 
	End Function

	Public Function GetProperty( _
		ByVal strName _
		)
 
		GetProperty = p_objFreeThreadedXmlDom.getProperty(strName) 
	End Function

	Public Function HasChildNodes()
		HasChildNodes = p_objFreeThreadedXmlDom.hasChildNodes()
	End Function

	Public Function InsertBefore( _
		ByVal objNewChild, _
		ByVal varRefChild _
		)
 
		Set InsertBefore = p_objFreeThreadedXmlDom.insertBefore(objNewChild, varRefChild) 
	End Function

	Public Function Load( _
		ByVal varXmlSource _
		)

		Load = p_objFreeThreadedXmlDom.load(varXmlSource)
	End Function

	Public Function LoadXml( _
		ByVal strXml _
		)

		LoadXml = p_objFreeThreadedXmlDom.loadXML(strXml)
	End Function

	Public Function NodeFromID( _
		ByVal varIdString _
		)
 
		Set NodeFromID = p_objFreeThreadedXmlDom.nodeFromID(varIdString) 
	End Function
  
	Public Function RemoveChild( _
		ByVal objChildNode _
		)

		Set RemoveChild = p_objFreeThreadedXmlDom.removeChild(objChildNode)
	End Function

	Public Function ReplaceChild( _
		ByVal objNewChild, _
		ByVal objOldChild _
		)

		Set ReplaceChild = p_objFreeThreadedXmlDom.replaceChild(objNewChild, objOldChild)
	End Function

	Public Sub Save( _
		ByVal varDestination _
		)
 
		p_objFreeThreadedXmlDom.save varDestination
	End Sub

	Public Function SelectNodes( _
		ByVal strQueryString _
		)

		Set SelectNodes = p_objFreeThreadedXmlDom.selectNodes(strQueryString)
	End Function

	Public Function SelectSingleNode( _
		ByVal strQueryString _
		)

		Set SelectSingleNode = p_objFreeThreadedXmlDom.selectSingleNode(strQueryString)
	End Function
 
	Public Sub SetProperty( _
		ByVal strName, _
		ByVal varValue _
		)
 
		p_objFreeThreadedXmlDom.setProperty strName, varValue
	End Sub

	Public Function TransformNode( _
		ByVal objStylesheet _
		)

		TransformNode = p_objFreeThreadedXmlDom.transformNode(objStylesheet)
	End Function

	Public Sub TransformNodeToObject( _
		ByVal objStylesheet, _
		ByVal varOutputObject _
		)
 
		p_objFreeThreadedXmlDom.transformNodeToObject objStylesheet, varOutputObject
	End Sub

	Public Function Validate()
		Set Validate = p_objFreeThreadedXmlDom.validate()
	End Function

	Private Sub Class_Terminate()
		Set p_objFreeThreadedXmlDom = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_XML_Microsoft_FreeThreadedXmlDom.vbs" Then

End If