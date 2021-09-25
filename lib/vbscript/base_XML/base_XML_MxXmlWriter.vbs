Option Explicit

Class base_XML_MxXmlWriter
	Private p_objMxXmlWriter

	Private Sub Class_Initialize()
		Set p_objMxXmlWriter = CreateObject("MSXML2.MXXMLWriter.6.0")
	End Sub


	' Properties


	Public Property Get ByteOrderMark()
		ByteOrderMark = p_objMxXmlWriter.byteOrderMark
	End Property

	Public Property Let ByteOrderMark( _
		ByVal blnByteOrderMark _
		)

		p_objMxXmlWriter.byteOrderMark = blnByteOrderMark
	End Property
 
	Public Property Get DisableOutputEscaping()
		DisableOutputEscaping = p_objMxXmlWriter.disableOutputEscaping
	End Property

	Public Property Let DisableOutputEscaping( _
		ByVal blnDisableOutputEscaping _
		)
 
		p_objMxXmlWriter.disableOutputEscaping = blnDisableOutputEscaping
	End Property

	Public Property Get Encoding()
		Encoding = p_objMxXmlWriter.encoding
	End Property

	Public Property Let Encoding( _
		ByVal strEncoding _
		)
 
		p_objMxXmlWriter.encoding = strEncoding
	End Property

	Public Property Get Indent()
		Indent = p_objMxXmlWriter.indent
	End Property

	Public Property Let Indent( _
		ByVal blnIndent _
		)
 
		p_objMxXmlWriter.indent = blnIndent
	End Property

	Public Property Get OmitXmlDeclaration()
		OmitXmlDeclaration = p_objMxXmlWriter.omitXMLDeclaration 
	End Property

	Public Property Let OmitXmlDeclaration( _
		ByVal blnOmitXmlDeclaration _
		)
 
		p_objMxXmlWriter.omitXMLDeclaration = blnOmitXmlDeclaration
	End Property

	Public Property Get Output()
		If IsObject(p_objMxXmlWriter.output) Then
			Set Output = p_objMxXmlWriter.output
		Else
			Output = p_objMxXmlWriter.output
		End If
	End Property

	Public Property Let Output( _
		ByVal varOutput _
		)
 
		p_objMxXmlWriter.output = varOutput
	End Property

	Public Property Set Output( _
		ByVal varOutput _
		)
 
		Set p_objMxXmlWriter.output = varOutput
	End Property

	Public Property Get Standalone()
		Standalone = p_objMxXmlWriter.standalone
	End Property

	Public Property Let Standalone( _
		ByVal blnStandalone _
		)
 
		p_objMxXmlWriter.standalone = blnStandalone
	End Property

	Public Property Get Version()
		Version = p_objMxXmlWriter.version
	End Property

	Public Property Let Version( _
		ByVal strVersion _
		)
 
		p_objMxXmlWriter.version = strVersion
	End Property


	' Methods


	Public Sub flush()
		p_objMxXmlWriter.flush
	End Sub

	Private Sub Class_Terminate()
		Set p_objMxXmlWriter = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_XML_MxXmlWriter.vbs" Then

End If