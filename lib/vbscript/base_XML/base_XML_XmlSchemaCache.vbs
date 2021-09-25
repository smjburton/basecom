Option Explicit

Class base_XML_XmlSchemaCache
	Private p_objXmlSchemaCache

	Private Sub Class_Initialize()
		Set p_objXmlSchemaCache = CreateObject("MSXML2.XMLSchemaCache.6.0")
	End Sub


	' Properties


	Public Property Get Length()
		p_objXmlSchemaCache.length
	End Property

	Public Default Property Get NamespaceUri( _
		ByVal lngIndex _
		)

		p_objXmlSchemaCache.namespaceURI
	End Property

	Public Property Get ValidateOnLoad()
		p_objXmlSchemaCache.validateOnLoad 
	End Property

	Public Property Let ValidateOnLoad( _
		ByVal blnValidateOnLoad _
		)

		p_objXmlSchemaCache.validateOnLoad 
	End Property


	' Methods


	Public Sub Add( _
		ByVal strNamespaceUri, _
		ByVal varSchema _
		)

		p_objXmlSchemaCache.add strNamespaceUri, varSchema
	End Sub
 
	Public Sub AddCollection( _
		ByVal objOtherCollection _
		)
 
		p_objXmlSchemaCache.addCollection objOtherCollection
	End Sub
 
	Public Function GetSchemaByNamespaceUri( _
		ByVal strNamespaceUri _
		)

		Set GetSchemaByNamespaceUri = p_objXmlSchemaCache.get(strNamespaceUri)
	End Function

	Public Function GetDeclaration( _
		ByVal objNode _
		)

		Set GetDeclaration = p_objXmlSchemaCache.getDeclaration(objNode)
	End Function

	Public Function GetSchema( _
		ByVal strNamespaceUri _
		)

		Set GetSchema = p_objXmlSchemaCache.getSchemaGetSchema
	End Function

	Public Sub Remove( _
		ByVal strNamespaceUri _
		)

		p_objXmlSchemaCache.remove strNamespaceUri
	End Sub
 
	Public Sub Validate()
		p_objXmlSchemaCache.validate
	End Sub
 
	Private Sub Class_Terminate()
		Set p_objXmlSchemaCache = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_XML_XmlSchemaCache.vbs" Then

End If