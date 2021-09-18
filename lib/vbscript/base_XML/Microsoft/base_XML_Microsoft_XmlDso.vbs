Option Explicit

Class base_XML_Microsoft_XmlDso
	Private p_objXmlDso

	Private Sub Class_Initialize()
		Set p_objXmlDso = CreateObject("Microsoft.XMLDSO.1.0")
	End Sub


	' Properties


	Public Property Get JavaDsoCompatible()
		p_objXmlDso.JavaDSOCompatible 
	End Property

	Public Property Let JavaDsoCompatible( _
		ByVal lngJavaDsoCompatible _
		)

		p_objXmlDso.JavaDSOCompatible = lngJavaDsoCompatible
	End Property

	Public Property Get ReadyState()
		ReadyState = p_objXmlDso.readyState 
	End Property

	Public Property Get XmlDocument()
		Set XmlDocument = p_objXmlDso.XMLDocument 
	End Property

	Public Property Set XmlDocument( _
		ByVal objXmlDocument _
		)

		Set p_objXmlDso.XMLDocument = objXmlDocument
	End Property

	Private Sub Class_Terminate()
		Set p_objXmlDso = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_XML_Microsoft_XmlDso.vbs" Then

End If