Option Explicit

Class base_XML_XmlSchemaCache
	Private p_objXmlSchemaCache

	Private Sub Class_Initialize()
		Set p_objXmlSchemaCache = CreateObject("MSXML2.XSLTemplate.6.0")
	End Sub


	' Properties


	Public Property Get Stylesheet()
		Set Stylesheet = p_objXmlSchemaCache.stylesheet
	End Property

	Public Property Set Stylesheet( _
		ByVal objStylesheet _
		)

		Set p_objXmlSchemaCache.stylesheet = objStylesheet
	End Property


	' Methods


	Public Function CreateProcessor()
		Set CreateProcessor = p_objXmlSchemaCache.createProcessor()
	End Function

	Private Sub Class_Terminate()
		Set p_objXmlSchemaCache = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_XML_XmlSchemaCache.vbs" Then

End If