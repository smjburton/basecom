Option Explicit

Class base_XML_MxHtmlWriter
	Private p_objMxHtmlWriter
	
	Private Sub Class_Initialize()
		Set p_objMxHtmlWriter = CreateObject("Msxml2.MXHTMLWriter.6.0")
	End Sub


	' Properties


	Public Property Get ByteOrderMark()
		ByteOrderMark = p_objMxHtmlWriter.byteOrderMark
	End Property

	Public Property Let ByteOrderMark( _
		ByVal blnByteOrderMark _
		)

		p_objMxHtmlWriter.byteOrderMark = blnByteOrderMark
	End Property

	Public Property Get DisableOutputEscaping()
		DisableOutputEscaping = p_objMxHtmlWriter.disableOutputEscaping
	End Property

	Public Property Let DisableOutputEscaping( _
		ByVal blnDisableOutputEscaping _
		)

		p_objMxHtmlWriter.disableOutputEscaping = blnDisableOutputEscaping
	End Property

	Public Property Get Encoding()
		Encoding = p_objMxHtmlWriter.encoding
	End Property

	Public Property Let Encoding( _
		ByVal strEncoding _
		)

		p_objMxHtmlWriter.encoding = strEncoding
	End Property

	Public Property Get Indent()
		Indent = p_objMxHtmlWriter.indent
	End Property

	Public Property Let Indent( _
		ByVal blnIndent _
		)

		p_objMxHtmlWriter.indent = blnIndent
	End Property

	Public Property Get OmitXmlDeclaration()
		OmitXmlDeclaration = p_objMxHtmlWriter.omitXMLDeclaration
	End Property

	Public Property Let OmitXmlDeclaration( _
		ByVal blnOmitXmlDeclaration _
		)
 
		p_objMxHtmlWriter.omitXMLDeclaration = blnOmitXmlDeclaration
	End Property

	Public Property Get Output()
		If IsObject(p_objMxHtmlWriter.output) Then
			Set Output = p_objMxHtmlWriter.output
		Else
			Output = p_objMxHtmlWriter.output
		End If
	End Property

	Public Property Let Output( _
		ByVal varOutput _
		)
 
		p_objMxHtmlWriter.output = varOutput
	End Property

	Public Property Set Output( _
		ByVal varOutput _
		)
 
		Set p_objMxHtmlWriter.output = varOutput
	End Property

	Public Property Get Standalone()
		Standalone = p_objMxHtmlWriter.standalone
	End Property

	Public Property Let Standalone( _
		ByVal blnStandalone _
		)

		p_objMxHtmlWriter.standalone = blnStandalone
	End Property

	Public Property Get Version()
		Version = p_objMxHtmlWriter.version
	End Property

	Public Property Let Version( _
		ByVal strVersion _
		)

		p_objMxHtmlWriter.version = strVersion
	End Property


	' Methods


	Public Sub Flush()
		p_objMxHtmlWriter.flush
	End Sub

	Private Sub Class_Terminate()
		Set p_objMxHtmlWriter = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_XML_MxHtmlWriter.vbs" Then

End If