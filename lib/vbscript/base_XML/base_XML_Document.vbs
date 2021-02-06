Option Explicit

Class base_XML_Document
	Private p_XmlDocument

	Private Sub Class_Initialize()
		Set p_XmlDocument = CreateObject("MSXML2.DOMDocument")
	End Sub

	
	' Properties


	Public Property Get Async()
		Async = p_XmlDocument.Async
	End Property

	Public Property Let Async(blnAsync)
		p_XmlDocument.Async = blnAsync
	End Property

	Public Property Get Attributes() 
		Set Attributes = p_XmlDocument.Attributes
	End Property

	Public Property Get BaseName()  
		BaseName = p_XmlDocument.BaseName
	End Property
 
	Public Property Get ChildNodes()
		Set ChildNodes = p_XmlDocument.ChildNodes
	End Property
 
	Public Property Get DataType()
		If IsObject(p_XmlDocument.DataType) Then
			Set DataType = p_XmlDocument.DataType
		Else
			DataType = p_XmlDocument.DataType
		End If
	End Property

	Public Property Let DataType(varDataType)
		p_XmlDocument.DataType = varDataType
	End Property

	Public Property Set DataType(varDataType)
		Set p_XmlDocument.DataType = varDataType
	End Property
 
	Public Property Get Definition()
		Set Definition = p_XmlDocument.Definition
	End Property

	Public Property Get DocType()
		Set DocType = p_XmlDocument.DocType
	End Property
 
	Public Property Get DocumentElement() 
		Set DocumentElement = p_XmlDocument.DocumentElement
	End Property

	Public Property Set DocumentElement(objIXMLDOMElement)
		Set p_XmlDocument.DocumentElement = objIXMLDOMElement
	End Property

	Public Property Get FirstChild() 
		Set FirstChild = p_XmlDocument.FirstChild
	End Property

	Public Property Get Implementation()
		Set Implementation = p_XmlDocument.Implementation
	End Property

	Public Property Get LastChild()
		Set LastChild = p_XmlDocument.LastChild
	End Property

	Public Property Get Namespaces()
		Set Namespaces = p_XmlDocument.Namespaces
	End Property

	Public Property Get NamespaceURI() 
		Namespaces = p_XmlDocument.Namespaces
	End Property

	Public Property Get NextSibling() 
		Set NextSibling = p_XmlDocument.NextSibling
	End Property

	Public Property Get NodeName()
		NodeName = p_XmlDocument.NodeName
	End Property

	Public Property Get NodeType()
		Set NodeType = p_XmlDocument.NodeType
	End Property

	Public Property Get NodeTypedValue()
		If IsObject(p_XmlDocument.NodeTypedValue) Then
			Set NodeTypedValue = p_XmlDocument.NodeTypedValue
		Else
			NodeTypedValue = p_XmlDocument.NodeTypedValue
		End If
	End Property

	Public Property Let NodeTypedValue(varNodeTypedValue)
		p_XmlDocument.NodeTypedValue = varNodeTypedValue
	End Property

	Public Property Set NodeTypedValue(varNodeTypedValue)
		Set p_XmlDocument.NodeTypedValue = varNodeTypedValue
	End Property
 
	Public Property Get NodeTypeString()
		NodeTypedValue = p_XmlDocument.NodeTypedValue
	End Property

	Public Property Get NodeValue()
		If IsObject(p_XmlDocument.NodeValue) Then
			Set NodeValue = p_XmlDocument.NodeValue
		Else
			NodeValue = p_XmlDocument.NodeValue
		End If
	End Property

	Public Property Let NodeValue(varNodeValue)
		p_XmlDocument.NodeValue = varNodeValue
	End Property

	Public Property Set NodeValue(varNodeValue)
		Set p_XmlDocument.NodeValue = varNodeValue
	End Property

	Public Property Get OnDataAvailable() 
		If IsObject(p_XmlDocument.OnDataAvailable) Then
			Set OnDataAvailable = p_XmlDocument.OnDataAvailable
		Else
			OnDataAvailable = p_XmlDocument.OnDataAvailable
		End If
	End Property

	Public Property Let OnDataAvailable(varOnDataAvailable) 
		p_XmlDocument.OnDataAvailable = varOnDataAvailable
	End Property

	Public Property Set OnDataAvailable(varOnDataAvailable)
		Set p_XmlDocument.OnDataAvailable = varOnDataAvailable
	End Property
 
	Public Property Get OnReadyStateChange() 
		If IsObject(p_XmlDocument.OnReadyStateChange) Then
			Set OnReadyStateChange = p_XmlDocument.OnReadyStateChange
		Else
			OnReadyStateChange = p_XmlDocument.OnReadyStateChange
		End If
	End Property

	Public Property Let OnReadyStateChange(varOnReadyStateChange)
		p_XmlDocument.OnReadyStateChange = varOnReadyStateChange
	End Property

	Public Property Set OnReadyStateChange(varOnReadyStateChange)
		Set p_XmlDocument.OnReadyStateChange = varOnReadyStateChange
	End Property

	Public Property Get OnTransformNode() 
		If IsObject(p_XmlDocument.OnTransformNode) Then
			Set OnTransformNode = p_XmlDocument.OnTransformNode
		Else
			OnTransformNode = p_XmlDocument.OnTransformNode
		End If
	End Property

	Public Property Let OnTransformNode(varOnTransformNode)
		p_XmlDocument.OnTransformNode = varOnTransformNode
	End Property

	Public Property Set OnTransformNode(varOnTransformNode)
		Set p_XmlDocument.OnTransformNode = varOnTransformNode
	End Property

	Public Property Get OwnerDocument()
		Set OwnerDocument = p_XmlDocument.OwnerDocument
	End Property

	Public Property Get ParentNode()
		Set ParentNode = p_XmlDocument.ParentNode
	End Property

	Public Property Get Parsed()
		Parsed = p_XmlDocument.Parsed
	End Property

	Public Property Get ParseError()
		Set ParseError = p_XmlDocument.ParseError
	End Property

	Public Property Get Prefix()
		Prefix = p_XmlDocument.Prefix
	End Property

	Public Property Get PreserveWhiteSpace() 
		PreserveWhiteSpace = p_XmlDocument.PreserveWhiteSpace
	End Property

	Public Property Let PreserveWhiteSpace(blnPreserveWhiteSpace)
		p_XmlDocument.PreserveWhiteSpace = blnPreserveWhiteSpace
	End Property

	Public Property Get PreviousSibling()
		Set PreviousSibling = p_XmlDocument.PreviousSibling
	End Property

	Public Property Get ReadyState() 
		PreviousSibling = p_XmlDocument.PreviousSibling
	End Property

	Public Property Get ResolveExternals()
		ResolveExternals = p_XmlDocument.ResolveExternals
	End Property

	Public Property Let ResolveExternals(blnResolveExternals)
		p_XmlDocument.ResolveExternals = blnResolveExternals
	End Property
 
	Public Property Get Schemas()
		If IsObject(p_XmlDocument.Schemas) Then
			Set Schemas = p_XmlDocument.Schemas
		Else
			Schemas = p_XmlDocument.Schemas
		End If
	End Property

	Public Property Let Schemas(varSchemas)
		p_XmlDocument.Schemas = varSchemas
	End Property

	Public Property Set Schemas(varSchemas)
		Set p_XmlDocument.Schemas = varSchemas
	End Property

	Public Property Get Specified()
		Specified = p_XmlDocument.Specified
	End Property

	Public Property Get Text() 
		Text = p_XmlDocument.Text
	End Property

	Public Property Let Text(strText)
		p_XmlDocument.Text = strText
	End Property

	Public Property Get URL() 
		URL = p_XmlDocument.URL
	End Property

	Public Property Get ValidateOnParse()
		ValidateOnParse = p_XmlDocument.ValidateOnParse
	End Property

	Public Property Let ValidateOnParse(blnValidateOnParse)
		p_XmlDocument.ValidateOnParse = blnValidateOnParse
	End Property

	Public Property Get XML()
		XML = p_XmlDocument.XML
	End Property


	' Methods


	Public Sub Abort()
		p_XmlDocument.Abort
	End Sub

	Public Function AppendChild(objNewChild)
		Set AppendChild = p_XmlDocument.AppendChild(objNewChild)
	End Function

	Public Function CloneNode(blnDeep)
		Set CloneNode = p_XmlDocument.CloneNode(blnDeep)
	End Function

	Public Function CreateAttribute(strName) 
		Set CreateAttribute = p_XmlDocument.CreateAttribute(strName) 
	End Function

	Public Function CreateCDATASection(strData)
		Set CreateCDATASection = p_XmlDocument.CreateCDATASection(strData)
	End Function
   
	Public Function CreateComment(strData) 
		Set CreateComment = p_XmlDocument.CreateComment(strData) 
	End Function

	Public Function CreateDocumentFragment() 
		Set CreateDocumentFragment = p_XmlDocument.CreateDocumentFragment()
	End Function

	Public Function CreateElement(strTagName)
		Set CreateElement = p_XmlDocument.CreateElement(strTagName)
	End Function
 
	Public Function CreateEntityReference(strName)
		Set CreateEntityReference = p_XmlDocument.CreateEntityReference(strName)
	End Function
 
	Public Function CreateNode(varType, strName, strNamespaceUri)
		Set CreateNode = p_XmlDocument.CreateNode(varType, strName, strNamespaceUri)
	End Function
 
	Public Function CreateProcessingInstruction(strTarget, strData)
		Set CreateProcessingInstruction = p_XmlDocument.CreateProcessingInstruction(strTarget, strData)
	End Function

	Public Function CreateTextNode(strData)
		Set CreateTextNode = p_XmlDocument.CreateTextNode(strData)
	End Function

	Public Function GetElementsByTagName(strTagName)
		Set GetElementsByTagName = p_XmlDocument.GetElementsByTagName(strTagName)
	End Function

	Public Function GetProperty(strName) 
		GetProperty = p_XmlDocument.GetProperty(strName)
	End Function

	Public Function HasChildNodes()
		HasChildNodes = p_XmlDocument.HasChildNodes()
	End Function

	Public Function ImportNode(objNode, blnDeep)
		Set ImportNode = p_XmlDocument.ImportNode(objNode, blnDeep)
	End Function

	Public Function InsertBefore(objNewChild, varRefChild) 
		Set InsertBefore = p_XmlDocument.InsertBefore(objNewChild, varRefChild)
	End Function

	Public Function Load(strXmlSource) 
		Load = p_XmlDocument.Load(strXmlSource) 
	End Function

	Public Function LoadXML(strXml) 
		LoadXML = p_XmlDocument.LoadXML(strXml) 
	End Function

	Public Function NodeFromID(strId)
		Set NodeFromID = p_XmlDocument.NodeFromID(strId)
	End Function

	Public Function RemoveChild(objChildNode) 
		Set RemoveChild = p_XmlDocument.RemoveChild(objChildNode) 
	End Function

	Public Function ReplaceChild(objNewChild, objOldChild)
		Set ReplaceChild = p_XmlDocument.ReplaceChild(objNewChild, objOldChild)
	End Function
 
	Public Sub Save(varDestination) 
		p_XmlDocument.Save varDestination
	End Sub

	Public Function SelectNodes(strQuery)
		Set SelectNodes = p_XmlDocument.SelectNodes(strQuery)
	End Function

	Public Function SelectSingleNode(strQuery) 
		Set SelectSingleNode = p_XmlDocument.SelectSingleNode(strQuery)
	End Function

	Public Sub SetProperty(strName, varValue) 
		p_XmlDocument.SetProperty strName, varValue
	End Sub

	Public Function TransformNode(objStyleSheet)
		TransformNode = p_XmlDocument.TransformNode(objStyleSheet)
	End Function

	Public Sub TransformNodeToObject(objStyleSheet, varOutputObject)
		p_XmlDocument.TransformNodeToObject objStyleSheet, varOutputObject
	End Sub
 
	Public Function Validate()
		Set Validate = p_XmlDocument.Validate
	End Function

	Public Function ValidateNode(objNode)
		Set ValidateNode = p_XmlDocument.ValidateNode(objNode)
	End Function

	Private Sub Class_Terminate()
		Set p_XmlDocument = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_XML_Document.vbs" Then

End If