Option Explicit

Class base_XML_SaxAttributes
	Private p_objSaxAttributes

	Private Sub Class_Initialize()
		Set p_objSaxAttributes = CreateObject("MSXML2.SAXAttributes.6.0")
	End Sub


	' Methods


	Public Sub AddAttribute( _
		ByVal strUri, _
		ByVal strLocalName, _
		ByVal strQName, _
		ByVal strType, _
		ByVal strValue _
		)

		p_objSaxAttributes.addAttribute strUri, strLocalName, strQName, strType, strValue
	End Sub

	Public Sub AddAttributeFromIndex( _
		ByVal varAtts, _
		ByVal lngIndex)
		p_objSaxAttributes.addAttributeFromIndex strUri, strLocalName, strQName, strType, strValue
	End Sub

	Public Sub Clear()
		p_objSaxAttributes.clear
	End Sub

	Public Sub RemoveAttribute( _
		ByVal lngIndex _
		)

		p_objSaxAttributes.removeAttribute lngIndex
	End Sub
 
	Public Sub SetAttribute( _
		ByVal lngIndex, _
		ByVal strUri, _
		ByVal strLocalName, _
		ByVal strQName, _
		ByVal strType, _
		ByVal strValue _
		)

		p_objSaxAttributes.setAttribute lngIndex, strUri, strLocalName, strQName, strType, strValue
	End Sub
 
	Public Sub SetAttributes( _
		ByVal varAtts _
		)

		p_objSaxAttributes.setAttributes varAtts
	End Sub
 
	Public Sub SetLocalName( _
		ByVal lngIndex, _
		ByVal strLocalName _
		)

		p_objSaxAttributes.setLocalName lngIndex, strLocalName
	End Sub
 
	Public Sub SetQName( _
		ByVal lngIndex, _
		ByVal strQName _
		)

		p_objSaxAttributes.setQName lngIndex, strQName
	End Sub
 
	Public Sub SetType( _
		ByVal lngIndex, _
		ByVal strType _
		)

		p_objSaxAttributes.setType lngIndex, strType
	End Sub
 
	Public Sub SetUri( _
		ByVal lngIndex, _
		ByVal strUri _
		)

		p_objSaxAttributes.setURI lngIndex, strUri
	End Sub
 
	Public Sub SetValue( _
		ByVal lngIndex, _
		ByVal strValue _
		)
 
		p_objSaxAttributes.setValue lngIndex, strValue
	End Sub

	Private Sub Class_Terminate()
		Set p_objSaxAttributes = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_XML_SaxAttributes.vbs" Then

End If