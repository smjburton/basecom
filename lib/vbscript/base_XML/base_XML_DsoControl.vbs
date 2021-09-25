Option Explicit

Class base_XML_DsoControl
	Private p_objDsoControl

	Private Sub Class_Initialize()
		Set p_objDsoControl = CreateObject("MSXML2.DSOControl.4.0")
	End Sub

	Public Property Get JavaDsoCompatible()
		JavaDsoCompatible = p_objDsoControl.JavaDSOCompatible 
	End Property

	Public Property Let JavaDsoCompatible( _
		ByVal lngJavaDsoCompatible _
		)

		p_objDsoControl.JavaDSOCompatible = lngJavaDsoCompatible
	End Property

	Public Property Get ReadyState()
		ReadyState = p_objDsoControl.readyState 
	End Property

	Public Property Get XmlDocument()
		Set XmlDocument = p_objDsoControl.XMLDocument 	
	End Property

	Public Property Set XmlDocument( _
		ByVal objXmlDocument _
		)

		Set p_objDsoControl.XMLDocument = objXmlDocument  
	End Property

	Private Sub Class_Terminate()
		Set p_objDsoControl = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_XML_DsoControl.vbs" Then

End If