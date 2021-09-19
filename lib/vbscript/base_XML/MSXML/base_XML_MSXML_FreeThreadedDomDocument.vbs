Option Explicit

Class base_XML_MSXML_FreeThreadedDomDocument
	Private p_objFreeThreadedDomDocument

	Private Sub Class_Initialize()
		Set p_objFreeThreadedDomDocument = CreateObject("MSXML.FreeThreadedDOMDocument")
	End Sub


	' Properties


	Public Property Get Async()
		Async = p_objFreeThreadedDomDocument.async 
	End Property

	Public Property Let Async( _
		ByVal blnAsync _
		)

		p_objFreeThreadedDomDocument.async = blnAsync
	End Property

	Public Property Get Attributes()
		Set Attributes = p_objFreeThreadedDomDocument.attributes 
	End Property

	Public Property Get BaseName()
		BaseName = p_objFreeThreadedDomDocument.baseName 
	End Property
 
	Public Property Get ChildNodes()
		Set ChildNodes = p_objFreeThreadedDomDocument.childNodes 
	End Property

	Public Property Get DataType()
		If IsObject(p_objFreeThreadedDomDocument.dataType) Then
			Set DataType = p_objFreeThreadedDomDocument.dataType
		Else
			DataType = p_objFreeThreadedDomDocument.dataType
		End If
	End Property

	Public Property Let DataType( _
		ByVal varDataType _
		)

		p_objFreeThreadedDomDocument.dataType = varDataType
	End Property

	Public Property Set DataType( _
		ByVal varDataType _
		)

		Set p_objFreeThreadedDomDocument.dataType = varDataType 
	End Property

	Public Property Get Definition()
		Set Definition = p_objFreeThreadedDomDocument.definition 
	End Property
 
	Public Property Get Doctype()
		Set Doctype = p_objFreeThreadedDomDocument.doctype 
	End Property

	Public Property Get DocumentElement()
		Set DocumentElement = p_objFreeThreadedDomDocument.documentElement 
	End Property

	Public Property Set DocumentElement( _
		ByVal objDocumentElement _
		)

		Set p_objFreeThreadedDomDocument.documentElement = objDocumentElement
	End Property
 
	Public Property Get FirstChild()
		Set FirstChild = p_objFreeThreadedDomDocument.firstChild 
	End Property

	Public Property Get Implementation()
		Set Implementation = p_objFreeThreadedDomDocument.implementation 
	End Property

	Public Property Get LastChild()
		Set LastChild = p_objFreeThreadedDomDocument.lastChild 
	End Property

	Public Property Get Namespaces()
		Set Namespaces = p_objFreeThreadedDomDocument.namespaces 
	End Property

	Public Property Get NamespaceUri()
		NamespaceUri = p_objFreeThreadedDomDocument.namespaceURI 
	End Property

	Public Property Get NextSibling()
		Set NextSibling = p_objFreeThreadedDomDocument.nextSibling 
	End Property
 
	Public Property Get NodeName()
		NodeName = p_objFreeThreadedDomDocument.nodeName 
	End Property

	Public Property Get NodeType()
		Set NodeType = p_objFreeThreadedDomDocument.nodeType 
	End Property

	Public Property Get NodeTypedValue()
		If IsObject(p_objFreeThreadedDomDocument.nodeTypedValue) Then
			Set NodeTypedValue = p_objFreeThreadedDomDocument.nodeTypedValue
		Else
			NodeTypedValue = p_objFreeThreadedDomDocument.nodeTypedValue
		End If 
	End Property

	Public Property Let NodeTypedValue( _
		ByVal varNodeTypedValue _
		)

		p_objFreeThreadedDomDocument.nodeTypedValue = varNodeTypedValue
	End Property

	Public Property Set NodeTypedValue( _
		ByVal varNodeTypedValue _
		)

		Set p_objFreeThreadedDomDocument.nodeTypedValue = varNodeTypedValue 
	End Property

	Public Property Get NodeTypeString()
		NodeTypeString = p_objFreeThreadedDomDocument.nodeTypeString 
	End Property

	Public Property Get NodeValue()
		If IsObject(p_objFreeThreadedDomDocument.nodeValue) Then
			Set NodeValue = p_objFreeThreadedDomDocument.nodeValue
		Else
			NodeValue = p_objFreeThreadedDomDocument.nodeValue
		End If
	End Property

	Public Property Let NodeValue( _
		ByVal varNodeValue _
		)

		p_objFreeThreadedDomDocument.nodeValue = varNodeValue
	End Property

	Public Property Set NodeValue( _
		ByVal varNodeValue _
		)

		Set p_objFreeThreadedDomDocument.nodeValue = varNodeValue 
	End Property

	Public Property Get OnDataAvailable()
		If IsObject(p_objFreeThreadedDomDocument.ondataavailable) Then
			Set OnDataAvailable = p_objFreeThreadedDomDocument.ondataavailable
		Else
			OnDataAvailable = p_objFreeThreadedDomDocument.ondataavailable
		End If 
	End Property

	Public Property Let OnDataAvailable( _
		ByVal varOnDataAvailable _
		)

		p_objFreeThreadedDomDocument.ondataavailable = varOnDataAvailable
	End Property

	Public Property Set OnDataAvailable( _
		ByVal varOnDataAvailable _
		)

		Set p_objFreeThreadedDomDocument.ondataavailable = varOnDataAvailable
	End Property

	Public Property Get OnReadyStateChange()
		If IsObject(p_objFreeThreadedDomDocument.onreadystatechange) Then
			Set OnReadyStateChange = p_objFreeThreadedDomDocument.onreadystatechange
		Else
			OnReadyStateChange = p_objFreeThreadedDomDocument.onreadystatechange
		End If
	End Property

	Public Property Let OnReadyStateChange( _
		ByVal varOnReadyStateChange _
		)

		p_objFreeThreadedDomDocument.onreadystatechange = varOnReadyStateChange
	End Property

	Public Property Set OnReadyStateChange( _
		ByVal varOnReadyStateChange _
		)

		Set p_objFreeThreadedDomDocument.onreadystatechange = varOnReadyStateChange
	End Property

	Public Property Get OnTransformNode()
		If IsObject(p_objFreeThreadedDomDocument.ontransformnode) Then
			Set OnTransformNode = p_objFreeThreadedDomDocument.ontransformnode
		Else
			OnTransformNode = p_objFreeThreadedDomDocument.ontransformnode
		End If 
	End Property

	Public Property Let OnTransformNode( _
		ByVal varOnTransformNode _
		)

		p_objFreeThreadedDomDocument.ontransformnode = varOnTransformNode
	End Property

	Public Property Set OnTransformNode( _
		ByVal varOnTransformNode _
		)

		Set p_objFreeThreadedDomDocument.ontransformnode = varOnTransformNode
	End Property

	Public Property Get OwnerDocument()
		Set OwnerDocument = p_objFreeThreadedDomDocument.ownerDocument 
	End Property

	Public Property Get ParentNode()
		Set ParentNode = p_objFreeThreadedDomDocument.parentNode 
	End Property

	Public Property Get Parsed()
		Parsed = p_objFreeThreadedDomDocument.parsed 
	End Property

	Public Property Get ParseError()
		Set ParseError = p_objFreeThreadedDomDocument.parseError 
	End Property

	Public Property Get Prefix()
		Prefix = p_objFreeThreadedDomDocument.prefix 
	End Property

	Public Property Get PreserveWhiteSpace()
		PreserveWhiteSpace = p_objFreeThreadedDomDocument.preserveWhiteSpace 
	End Property

	Public Property Let PreserveWhiteSpace( _
		ByVal blnPreserveWhiteSpace _
		)

		p_objFreeThreadedDomDocument.preserveWhiteSpace = blnPreserveWhiteSpace
	End Property

	Public Property Get PreviousSibling()
		Set PreviousSibling = p_objFreeThreadedDomDocument.previousSibling 
	End Property

	Public Property Get ReadyState()
		ReadyState = p_objFreeThreadedDomDocument.readyState 
	End Property

	Public Property Get ResolveExternals()
		ResolveExternals = p_objFreeThreadedDomDocument.resolveExternals 
	End Property

	Public Property Let ResolveExternals( _
		ByVal blnResolveExternals _
		)

		p_objFreeThreadedDomDocument.resolveExternals = blnResolveExternals
	End Property

	Public Property Get Schemas()
		If IsObject(p_objFreeThreadedDomDocument.schemas) Then
			Set Schemas = p_objFreeThreadedDomDocument.schemas
		Else
			Schemas = p_objFreeThreadedDomDocument.schemas
		End If 
	End Property

	Public Property Let Schemas( _
		ByVal varSchemas _
		)

		p_objFreeThreadedDomDocument.schemas = varSchemas
	End Property

	Public Property Set Schemas( _
		ByVal varSchemas _
		)

		Set p_objFreeThreadedDomDocument.schemas = varSchemas
	End Property

	Public Property Get Specified()
		Specified = p_objFreeThreadedDomDocument.specified 
	End Property

	Public Property Get Text()
		Text = p_objFreeThreadedDomDocument.text 
	End Property

	Public Property Let Text( _
		ByVal strText _
		)

		p_objFreeThreadedDomDocument.text = strText
	End Property

	Public Property Get Url()
		Url = p_objFreeThreadedDomDocument.url 
	End Property

	Public Property Get ValidateOnParse()
		ValidateOnParse = p_objFreeThreadedDomDocument.validateOnParse 
	End Property

	Public Property Let ValidateOnParse( _
		ByVal blnValidateOnParse _
		)

		p_objFreeThreadedDomDocument.validateOnParse = blnValidateOnParse
	End Property

	Public Property Get Xml()
		Xml = p_objFreeThreadedDomDocument.xml 
	End Property


	' Methods


	Public Sub Abort()
		p_objFreeThreadedDomDocument.abort
	End Sub

	Public Function AppendChild( _
		ByVal objNewChild _
		)

		Set AppendChild = p_objFreeThreadedDomDocument.appendChild(objNewChild)
	End Function

	Public Function CloneNode( _
		ByVal blnDeep _
		)

		Set CloneNode = p_objFreeThreadedDomDocument.cloneNode(blnDeep)
	End Function

	Public Function CreateAttribute( _
		ByVal strName _
		)

		Set CreateAttribute = p_objFreeThreadedDomDocument.createAttribute(strName)
	End Function
 
	Public Function CreateCdataSection( _
		ByVal strData _
		)

		Set CreateCdataSection = p_objFreeThreadedDomDocument.createCDATASection(strData)
	End Function
 
	Public Function CreateComment( _
		ByVal strData _
		)

		Set CreateComment = p_objFreeThreadedDomDocument.createComment(strData)
	End Function

	Public Function CreateDocumentFragment()
		Set CreateDocumentFragment = p_objFreeThreadedDomDocument.createDocumentFragment()
	End Function
 
	Public Function CreateElement( _
		ByVal strTagName _
		)
 
		Set CreateElement = p_objFreeThreadedDomDocument.createElement(strTagName)
	End Function

	Public Function CreateEntityReference( _
		ByVal strName _
		)

		Set CreateEntityReference = p_objFreeThreadedDomDocument.createEntityReference(strName)
	End Function

	Public Function CreateNode( _
		ByVal varType, _
		ByVal strName, _
		ByVal strNamespaceUri _
		)

		Set CreateNode = p_objFreeThreadedDomDocument.createNode(varType, strName, strNamespaceUri)
	End Function

	Public Function CreateProcessingInstruction( _
		ByVal strTarget, _
		ByVal strData _
		)
 
		Set CreateProcessingInstruction = p_objFreeThreadedDomDocument.createProcessingInstruction(strTarget, strData)
	End Function

	Public Function CreateTextNode( _
		ByVal strData _
		)
 
		Set CreateTextNode = p_objFreeThreadedDomDocument.createTextNode(strData) 
	End Function

	Public Function GetElementsByTagName( _
		ByVal strTagName _
		)

		Set GetElementsByTagName = p_objFreeThreadedDomDocument.getElementsByTagName(strTagName)
	End Function

	Public Function GetProperty( _
		ByVal strName _
		)

		GetProperty = p_objFreeThreadedDomDocument.getProperty(strName)
	End Function

	Public Function HasChildNodes()
		HasChildNodes = p_objFreeThreadedDomDocument.hasChildNodes()
	End Function

	Public Function InsertBefore( _
		ByVal objNewChild, _
		ByVal objRefChild _
		)

		Set InsertBefore = p_objFreeThreadedDomDocument.insertBefore(objNewChild, objRefChild)
	End Function
 
	Public Function Load( _
		ByVal varXmlSource _
		)

		Load = p_objFreeThreadedDomDocument.load(varXmlSource)
	End Function

	Public Function LoadXml( _
		ByVal strXml _
		)

		LoadXml = p_objFreeThreadedDomDocument.loadXML(strXml)
	End Function

	Public Function NodeFromId( _
		ByVal strIdString _
		)
   
		Set NodeFromId = p_objFreeThreadedDomDocument.nodeFromID(strIdString)
	End Function
 
	Public Function RemoveChild( _
		ByVal objChildNode _
		)

		Set RemoveChild = p_objFreeThreadedDomDocument.removeChild(objChildNode)
	End Function

	Public Function ReplaceChild( _
		ByVal objNewChild, _
		ByVal objOldChild _
		)
 
		Set ReplaceChild = p_objFreeThreadedDomDocument.replaceChild(objNewChild, objOldChild)
	End Function

	Public Sub Save( _
		ByVal varDestination _
		)

		p_objFreeThreadedDomDocument.save varDestination
	End Sub

	Public Function SelectNodes( _
		ByVal strQueryString _
		)

		Set SelectNodes = p_objFreeThreadedDomDocument.selectNodes(strQueryString)
	End Function

	Public Function SelectSingleNode( _
		ByVal strQueryString _
		)

		Set SelectSingleNode = p_objFreeThreadedDomDocument.selectSingleNode(strQueryString)
	End Function

	Public Sub SetProperty( _
		ByVal strName, _
		ByVal varValue _
		)

		p_objFreeThreadedDomDocument.setProperty strName, varValue
	End Sub

	Public Function TransformNode( _
		ByVal objStylesheet _
		)

		TransformNode = p_objFreeThreadedDomDocument.transformNode(objStylesheet)
	End Function

	Public Sub TransformNodeToObject( _
		ByVal objStylesheet, _
		ByVal varOutputObject _
		)

		p_objFreeThreadedDomDocument.transformNodeToObject objStylesheet, varOutputObject
	End Sub

	Public Function Validate()
		Set Validate = p_objFreeThreadedDomDocument.validate()
	End Function

	Private Sub Class_Terminate()
		Set p_objFreeThreadedDomDocument = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_XML_MSXML_DomDocument.vbs" Then

End If