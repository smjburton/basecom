Option Explicit

Include "base_Data_Array"

Class base_XML_Parser
	Private p_objXmlDocument, _
		p_objSelection

	Private Sub Class_Initialize()
		Set p_objXmlDocument = CreateObject("MSXML2.DOMDocument")
		Set p_objSelection = New base_Data_Array
	End Sub


	' Properties


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


	Public Property Get Definition()
		Set Definition = p_objXmlDocument.Definition
	End Property

	Public Property Get DocType()
		Set DocType = p_objXmlDocument.DocType
	End Property
 
	Public Property Get DocumentElement() 
		Set DocumentElement = p_objXmlDocument.DocumentElement
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

	Public Property Get PreviousSibling()
		Set PreviousSibling = p_objXmlDocument.PreviousSibling
	End Property

	Public Property Get ReadyState() 
		ReadyState = p_objXmlDocument.ReadyState
	End Property

	Public Property Get Schemas()
		If IsObject(p_objXmlDocument.Schemas) Then
			Set Schemas = p_objXmlDocument.Schemas
		Else
			Schemas = p_objXmlDocument.Schemas
		End If
	End Property

	Public Property Get Specified()
		Specified = p_objXmlDocument.Specified
	End Property

	Public Property Get Attributes() 
		Set Attributes = p_objXmlDocument.Attributes
	End Property

	Public Property Get Text() 
		Text = p_objXmlDocument.Text
	End Property

	Public Property Get URL() 
		URL = p_objXmlDocument.URL
	End Property

	Public Property Get XML()
		XML = p_objXmlDocument.XML
	End Property

	' Methods



	Public Function GetElementsByTagName(strTagName)
		Set GetElementsByTagName = p_objXmlDocument.GetElementsByTagName(strTagName)
	End Function

	Public Function GetProperty(strName) 
		GetProperty = p_objXmlDocument.GetProperty(strName)
	End Function

	Public Function HasChildNodes()
		HasChildNodes = p_objXmlDocument.HasChildNodes()
	End Function


	Public Function NodeFromID(strId)
		Set NodeFromID = p_objXmlDocument.NodeFromID(strId)
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




	Public Sub Abort()
		p_objXmlDocument.Abort
	End Sub

	Public Sub FromString( _
		ByVal strXml _
		)

		p_objXmlDocument.LoadXML(strXml)
	End Sub

	Public Sub Load( _
		ByVal strXmlSource _
		)

		p_objXmlDocument.Load(strXmlSource)
	End Sub

	Public Sub Save( _
		ByVal varDestination _
		)

		p_objXmlDocument.Save varDestination
	End Sub

	Private Sub Class_Terminate()
		Set p_objXmlDocument = Nothing
		Set p_objSelection = Nothing
	End Sub
End Class

' Properties:

' Tag()
' Element(intIndex)
' Attribute(strAttr)
' Attributes()
' ID()
' ClassName()
' CurrentStyle()
' Style()
' NodeName()
' NodeType()
' NodeValue()
' Item(intIndex)
' Count()
' Text()
' HTML()
' Title()

' Preconfigured Selections:

' A()
' Body()
' Buttons()
' Comments()
' Divs()
' Forms()
' H1()
' H2()
' H3()
' H4()
' H5()
' H6()
' Head()
' Images()
' Inputs()
' Labels()
' LI()
' Links()
' Meta()
' OL()
' P()
' Scripts()
' Spans()
' Styles()
' StyleSheets()
' Tables()
' UL()

' Methods

' Traversal Methods

' Ascend(intLevel, intIndex)
' Descend(intLevel, intIndex)
' PreviousNode(intIndex)
' NextNode(intIndex)
' Offset(intLevel, intIndex)
' Ancestor(intLevel, intIndex)
' Descendant(intLevel, intIndex)
' Parents(intLevel)
' Child(intIndex)
' FirstChild()
' LastChild()
' Sibling(intIndex)
' Parent()
' Children()
' Siblings()
' NextSibling()
' PreviousSibling()
' Relative(intLevel, intIndex)

' Selection Methods

' All()
' First()
' Last()
' UpTo(varCondition)
' Slice(intStart, intEnd)
' Has(strAttr)
' Contains(varCondition)
' IsSelected(varCondition)
' Exclude(varCondition)
' Limit(intLimit)

' Retrieval Methods

' ElementsByName(strName)
' ElementsByTagName(strTagName)
' ElementsByClassName(strClassName)
' ElementsByAttribute(strAttr)
' ElementsByAttributeValue(strAttr, strValue)
' ElementByID(strID)
' Selector(strSelector)
' SelectorAll(strSelectors)
' Match(strType, strPattern, blnIgnoreCase, blnGlobal)
' MatchAll(strType, strPattern, blnIgnoreCase, blnGlobal)

' HTMLParser(varHTML)
' Load(strFile)
' FromString(strHTML)
' ToString()
' FromDocument(objHtmlDoc)
' ToDocument()
' SelectionFrom(objHtmlParser, objArr)
' Clear()
' ClearSelection()

If WScript.ScriptName = "base_XML_Parser.vbs" Then

End If