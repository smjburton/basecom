Option Explicit

Class base_XML_Document
	Private p_objXmlDocument

	Private Sub Class_Initialize()
		Set p_objXmlDocument = CreateObject("MSXML2.DOMDocument")
	End Sub

	
	' Properties


	Public Property Get Async()
		Async = p_objXmlDocument.Async
	End Property

	Public Property Let Async(blnAsync)
		p_objXmlDocument.Async = blnAsync
	End Property

	Public Property Get Attributes() 
		Set Attributes = p_objXmlDocument.Attributes
	End Property

	Public Property Get BaseName()  
		BaseName = p_objXmlDocument.BaseName
	End Property
 
	Public Property Get ChildNodes()
		Set ChildNodes = p_objXmlDocument.ChildNodes
	End Property
 
	Public Property Get DataType()
		If IsObject(p_objXmlDocument.DataType) Then
			Set DataType = p_objXmlDocument.DataType
		Else
			DataType = p_objXmlDocument.DataType
		End If
	End Property

	Public Property Let DataType(varDataType)
		p_objXmlDocument.DataType = varDataType
	End Property

	Public Property Set DataType(varDataType)
		Set p_objXmlDocument.DataType = varDataType
	End Property
 
	Public Property Get Definition()
		Set Definition = p_objXmlDocument.Definition
	End Property

	Public Property Get DocType()
		Set DocType = p_objXmlDocument.DocType
	End Property
 
	Public Property Get DocumentElement() 
		Set DocumentElement = p_objXmlDocument.DocumentElement
	End Property

	Public Property Set DocumentElement(objIXMLDOMElement)
		Set p_objXmlDocument.DocumentElement = objIXMLDOMElement
	End Property

	Public Property Get FirstChild() 
		Set FirstChild = p_objXmlDocument.FirstChild
	End Property

	Public Property Get Implementation()
		Set Implementation = p_objXmlDocument.Implementation
	End Property

	Public Property Get LastChild()
		Set LastChild = p_objXmlDocument.LastChild
	End Property

	Public Property Get Namespaces()
		Set Namespaces = p_objXmlDocument.Namespaces
	End Property

	Public Property Get NamespaceURI() 
		Namespaces = p_objXmlDocument.Namespaces
	End Property

	Public Property Get NextSibling() 
		Set NextSibling = p_objXmlDocument.NextSibling
	End Property

	Public Property Get NodeName()
		NodeName = p_objXmlDocument.NodeName
	End Property

	Public Property Get NodeType()
		Set NodeType = p_objXmlDocument.NodeType
	End Property

	Public Property Get NodeTypedValue()
		If IsObject(p_objXmlDocument.NodeTypedValue) Then
			Set NodeTypedValue = p_objXmlDocument.NodeTypedValue
		Else
			NodeTypedValue = p_objXmlDocument.NodeTypedValue
		End If
	End Property

	Public Property Let NodeTypedValue(varNodeTypedValue)
		p_objXmlDocument.NodeTypedValue = varNodeTypedValue
	End Property

	Public Property Set NodeTypedValue(varNodeTypedValue)
		Set p_objXmlDocument.NodeTypedValue = varNodeTypedValue
	End Property
 
	Public Property Get NodeTypeString()
		NodeTypeString = p_objXmlDocument.NodeTypeString
	End Property

	Public Property Get NodeValue()
		If IsObject(p_objXmlDocument.NodeValue) Then
			Set NodeValue = p_objXmlDocument.NodeValue
		Else
			NodeValue = p_objXmlDocument.NodeValue
		End If
	End Property

	Public Property Let NodeValue(varNodeValue)
		p_objXmlDocument.NodeValue = varNodeValue
	End Property

	Public Property Set NodeValue(varNodeValue)
		Set p_objXmlDocument.NodeValue = varNodeValue
	End Property

	Public Property Get OnDataAvailable() 
		If IsObject(p_objXmlDocument.OnDataAvailable) Then
			Set OnDataAvailable = p_objXmlDocument.OnDataAvailable
		Else
			OnDataAvailable = p_objXmlDocument.OnDataAvailable
		End If
	End Property

	Public Property Let OnDataAvailable(varOnDataAvailable) 
		p_objXmlDocument.OnDataAvailable = varOnDataAvailable
	End Property

	Public Property Set OnDataAvailable(varOnDataAvailable)
		Set p_objXmlDocument.OnDataAvailable = varOnDataAvailable
	End Property
 
	Public Property Get OnReadyStateChange() 
		If IsObject(p_objXmlDocument.OnReadyStateChange) Then
			Set OnReadyStateChange = p_objXmlDocument.OnReadyStateChange
		Else
			OnReadyStateChange = p_objXmlDocument.OnReadyStateChange
		End If
	End Property

	Public Property Let OnReadyStateChange(varOnReadyStateChange)
		p_objXmlDocument.OnReadyStateChange = varOnReadyStateChange
	End Property

	Public Property Set OnReadyStateChange(varOnReadyStateChange)
		Set p_objXmlDocument.OnReadyStateChange = varOnReadyStateChange
	End Property

	Public Property Get OnTransformNode() 
		If IsObject(p_objXmlDocument.OnTransformNode) Then
			Set OnTransformNode = p_objXmlDocument.OnTransformNode
		Else
			OnTransformNode = p_objXmlDocument.OnTransformNode
		End If
	End Property

	Public Property Let OnTransformNode(varOnTransformNode)
		p_objXmlDocument.OnTransformNode = varOnTransformNode
	End Property

	Public Property Set OnTransformNode(varOnTransformNode)
		Set p_objXmlDocument.OnTransformNode = varOnTransformNode
	End Property

	Public Property Get OwnerDocument()
		Set OwnerDocument = p_objXmlDocument.OwnerDocument
	End Property

	Public Property Get ParentNode()
		Set ParentNode = p_objXmlDocument.ParentNode
	End Property

	Public Property Get Parsed()
		Parsed = p_objXmlDocument.Parsed
	End Property

	Public Property Get ParseError()
		Set ParseError = p_objXmlDocument.ParseError
	End Property

	Public Property Get Prefix()
		Prefix = p_objXmlDocument.Prefix
	End Property

	Public Property Get PreserveWhiteSpace() 
		PreserveWhiteSpace = p_objXmlDocument.PreserveWhiteSpace
	End Property

	Public Property Let PreserveWhiteSpace(blnPreserveWhiteSpace)
		p_objXmlDocument.PreserveWhiteSpace = blnPreserveWhiteSpace
	End Property

	Public Property Get PreviousSibling()
		Set PreviousSibling = p_objXmlDocument.PreviousSibling
	End Property

	Public Property Get ReadyState() 
		ReadyState = p_objXmlDocument.ReadyState
	End Property

	Public Property Get ResolveExternals()
		ResolveExternals = p_objXmlDocument.ResolveExternals
	End Property

	Public Property Let ResolveExternals(blnResolveExternals)
		p_objXmlDocument.ResolveExternals = blnResolveExternals
	End Property
 
	Public Property Get Schemas()
		If IsObject(p_objXmlDocument.Schemas) Then
			Set Schemas = p_objXmlDocument.Schemas
		Else
			Schemas = p_objXmlDocument.Schemas
		End If
	End Property

	Public Property Let Schemas(varSchemas)
		p_objXmlDocument.Schemas = varSchemas
	End Property

	Public Property Set Schemas(varSchemas)
		Set p_objXmlDocument.Schemas = varSchemas
	End Property

	Public Property Get Specified()
		Specified = p_objXmlDocument.Specified
	End Property

	Public Property Get Text() 
		Text = p_objXmlDocument.Text
	End Property

	Public Property Let Text(strText)
		p_objXmlDocument.Text = strText
	End Property

	Public Property Get URL() 
		URL = p_objXmlDocument.URL
	End Property

	Public Property Get ValidateOnParse()
		ValidateOnParse = p_objXmlDocument.ValidateOnParse
	End Property

	Public Property Let ValidateOnParse(blnValidateOnParse)
		p_objXmlDocument.ValidateOnParse = blnValidateOnParse
	End Property

	Public Property Get XML()
		XML = p_objXmlDocument.XML
	End Property


	' Methods


	Public Sub Abort()
		p_objXmlDocument.Abort
	End Sub

	Public Function AppendChild(objNewChild)
		Set AppendChild = p_objXmlDocument.AppendChild(objNewChild)
	End Function

	Public Function CloneNode(blnDeep)
		Set CloneNode = p_objXmlDocument.CloneNode(blnDeep)
	End Function

	Public Function CreateAttribute(strName) 
		Set CreateAttribute = p_objXmlDocument.CreateAttribute(strName) 
	End Function

	Public Function CreateCDATASection(strData)
		Set CreateCDATASection = p_objXmlDocument.CreateCDATASection(strData)
	End Function
   
	Public Function CreateComment(strData) 
		Set CreateComment = p_objXmlDocument.CreateComment(strData) 
	End Function

	Public Function CreateDocumentFragment() 
		Set CreateDocumentFragment = p_objXmlDocument.CreateDocumentFragment()
	End Function

	Public Function CreateElement(strTagName)
		Set CreateElement = p_objXmlDocument.CreateElement(strTagName)
	End Function
 
	Public Function CreateEntityReference(strName)
		Set CreateEntityReference = p_objXmlDocument.CreateEntityReference(strName)
	End Function
 
	Public Function CreateNode(varType, strName, strNamespaceUri)
		Set CreateNode = p_objXmlDocument.CreateNode(varType, strName, strNamespaceUri)
	End Function
 
	Public Function CreateProcessingInstruction(strTarget, strData)
		Set CreateProcessingInstruction = p_objXmlDocument.CreateProcessingInstruction(strTarget, strData)
	End Function

	Public Function CreateTextNode(strData)
		Set CreateTextNode = p_objXmlDocument.CreateTextNode(strData)
	End Function

	Public Function GetElementsByTagName(strTagName)
		Set GetElementsByTagName = p_objXmlDocument.GetElementsByTagName(strTagName)
	End Function

	Public Function GetProperty(strName) 
		GetProperty = p_objXmlDocument.GetProperty(strName)
	End Function

	Public Function HasChildNodes()
		HasChildNodes = p_objXmlDocument.HasChildNodes()
	End Function

	Public Function ImportNode(objNode, blnDeep)
		Set ImportNode = p_objXmlDocument.ImportNode(objNode, blnDeep)
	End Function

	Public Function InsertBefore(objNewChild, varRefChild) 
		Set InsertBefore = p_objXmlDocument.InsertBefore(objNewChild, varRefChild)
	End Function

	Public Function Load(strXmlSource) 
		Load = p_objXmlDocument.Load(strXmlSource) 
	End Function

	Public Function LoadXML(strXml) 
		LoadXML = p_objXmlDocument.LoadXML(strXml) 
	End Function

	Public Function NodeFromID(strId)
		Set NodeFromID = p_objXmlDocument.NodeFromID(strId)
	End Function

	Public Function RemoveChild(objChildNode) 
		Set RemoveChild = p_objXmlDocument.RemoveChild(objChildNode) 
	End Function

	Public Function ReplaceChild(objNewChild, objOldChild)
		Set ReplaceChild = p_objXmlDocument.ReplaceChild(objNewChild, objOldChild)
	End Function
 
	Public Sub Save(varDestination) 
		p_objXmlDocument.Save varDestination
	End Sub

	Public Function SelectNodes(strQuery)
		Set SelectNodes = p_objXmlDocument.SelectNodes(strQuery)
	End Function

	Public Function SelectSingleNode(strQuery) 
		Set SelectSingleNode = p_objXmlDocument.SelectSingleNode(strQuery)
	End Function

	Public Sub SetProperty(strName, varValue) 
		p_objXmlDocument.SetProperty strName, varValue
	End Sub

	Public Function TransformNode(objStyleSheet)
		TransformNode = p_objXmlDocument.TransformNode(objStyleSheet)
	End Function

	Public Sub TransformNodeToObject(objStyleSheet, varOutputObject)
		p_objXmlDocument.TransformNodeToObject objStyleSheet, varOutputObject
	End Sub
 
	Public Function Validate()
		Set Validate = p_objXmlDocument.Validate
	End Function

	Public Function ValidateNode(objNode)
		Set ValidateNode = p_objXmlDocument.ValidateNode(objNode)
	End Function

	Private Sub Class_Terminate()
		Set p_objXmlDocument = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_XML_Document.vbs" Then

End If