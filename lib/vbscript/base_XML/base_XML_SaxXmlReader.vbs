Option Explicit

Class base_XML_SaxXmlReader
	Private p_objSaxXmlReader

	Private Sub Class_Initialize()
		Set p_objSaxXmlReader = CreateObject("MSXML2.SAXXMLReader.6.0")
	End Sub


	' Properties


	Public Property Get BaseUrl()
		p_objSaxXmlReader.baseURL
	End Property

	Public Property Let BaseUrl( _
		ByVal strBaseUrl _
		)

		p_objSaxXmlReader.baseURL
	End Property

	Public Property Get ContentHandler()
		Set p_objSaxXmlReader.contentHandler
	End Property

	Public Property Set ContentHandler( _
		ByVal objContentHandler _
		)

		Set p_objSaxXmlReader.contentHandler
	End Property

	Public Property Get DtdHandler()
		Set p_objSaxXmlReader.dtdHandler
	End Property

	Public Property Set DtdHandler( _
		ByVal objDtdHandler _
		)

		Set p_objSaxXmlReader.dtdHandler
	End Property

	Public Property Get EntityResolver()
		Set p_objSaxXmlReader.entityResolver
	End Property

	Public Property Set EntityResolver( _
		ByVal objEntityResolver _
		)
 
		Set p_objSaxXmlReader.entityResolver 
	End Property

	Public Property Get ErrorHandler()
		Set p_objSaxXmlReader.errorHandler
	End Property

	Public Property Set ErrorHandler( _
		ByVal objErrorHandler _
		)
 
		Set p_objSaxXmlReader.errorHandler
	End Property

	Public Property Get SecureBaseUrl()
		p_objSaxXmlReader.secureBaseURL
	End Property

	Public Property Let SecureBaseUrl( _
		ByVal strSecureBaseUrl _
		)

		p_objSaxXmlReader.secureBaseURL
	End Property


	' Methods


	Public Function GetFeature( _
		ByVal strName _
		)

		GetFeature = p_objSaxXmlReader.getFeature(strName)
	End Function

	Public Function GetProperty( _
		ByVal strName _
		)

		GetProperty = p_objSaxXmlReader.getProperty(strName)
	End Function

	Public Sub Parse( _
		ByVal varInput _
		)

		p_objSaxXmlReader.parse varInput
	End Sub

	Public Sub ParseUrl( _
		ByVal strUrl _
		)

		p_objSaxXmlReader.parseURL strUrl
	End Sub

	Public Sub PutFeature( _
		ByVal strName, _
		ByVal blnValue _
		)

		p_objSaxXmlReader.putFeature strName, blnValue
	End Sub

	Public Sub PutProperty( _
		ByVal strName, _
		ByVal blnValue varValue _
		)

		p_objSaxXmlReader.putProperty strName, varValue
	End Sub

	Private Sub Class_Terminate()
		Set p_objSaxXmlReader = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_XML_SaxXmlReader.vbs" Then

End If