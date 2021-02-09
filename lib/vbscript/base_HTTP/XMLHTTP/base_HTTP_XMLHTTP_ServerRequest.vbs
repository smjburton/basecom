Option Explicit

' “version-independent” ProgID for MSXML:
' Msxml2.ServerXMLHTTP = IServerXMLHTTPRequest object

' Msxml2.ServerXMLHTTP.3.0 = IServerXMLHTTPRequest object
' Msxml2.ServerXMLHTTP.4.0 = (?)
' Msxml2.ServerXMLHTTP.6.0 = IServerXMLHTTPRequest2 object

Class base_HTTP_XMLHTTP_ServerRequest
	Private p_XmlHttpServerReq

	Private Sub Class_Initialize()
		Set p_XmlHttpServerReq = CreateObject("MSXML2.ServerXMLHTTP")
	End Sub


	' Properties


	Public Property Get ReadyState()
		ReadyState = p_XmlHttpServerReq.ReadyState
	End Property

	Public Property Get ResponseBody()
		If IsObject(p_XmlHttpServerReq.ResponseBody) Then
			Set ResponseBody = p_XmlHttpServerReq.ResponseBody
		Else
			ResponseBody = p_XmlHttpServerReq.ResponseBody
		End If
	End Property

	Public Property Get ResponseStream()
		If IsObject(p_XmlHttpServerReq.ResponseStream) Then
			Set ResponseStream = p_XmlHttpServerReq.ResponseStream
		Else
			ResponseStream = p_XmlHttpServerReq.ResponseStream
		End If
	End Property

	Public Property Get ResponseText()
		ResponseText = p_XmlHttpServerReq.ResponseText
	End Property

	Public Property Get ResponseXML()
		Set ResponseXML = p_XmlHttpServerReq.ResponseXML
	End Property

	Public Property Get Status()
		Status = p_XmlHttpServerReq.Status
	End Property

	Public Property Get StatusText()
		StatusText = p_XmlHttpServerReq.StatusText
	End Property


	' Methods


	Public Sub Abort()
		p_XmlHttpServerReq.Abort
	End Sub

	Public Function GetAllResponseHeaders()
		GetAllResponseHeaders = p_XmlHttpServerReq.GetAllResponseHeaders()
	End Function

	Public Function GetOption(intSERVERXMLHTTP_OPTION)
		GetOption = p_XmlHttpServerReq.GetOption(intSERVERXMLHTTP_OPTION)
	End Function

	Public Function GetResponseHeader(strHeader)
		GetResponseHeader = p_XmlHttpServerReq.GetResponseHeader(strHeader)
	End Function

	Public Sub Open(strMethod, strUrl) ' Optional params: [varAsync], [bstrUser], [bstrPassword]
		p_XmlHttpServerReq.Open strMethod, strUrl
	End Sub

	Public Sub Send() ' Optional params: [varBody]
		p_XmlHttpServerReq.Send
	End Sub

	Public Sub SetOption(intSERVERXMLHTTP_OPTION, varValue)
		p_XmlHttpServerReq.SetOption intSERVERXMLHTTP_OPTION, varValue
	End Sub

	Public Sub SetProxy(intSXH_PROXY_SETTING) ' Optional params: [varProxyServer], [varBypassList]
		p_XmlHttpServerReq.SetProxy intSXH_PROXY_SETTING
	End Sub

	Public Sub SetProxyCredentials(strUserName, strPassword)
		p_XmlHttpServerReq.SetProxyCredentials strUserName, strPassword
	End Sub

	Public Sub SetRequestHeader(strHeader, strValue)
		p_XmlHttpServerReq.SetRequestHeader strHeader, strValue
	End Sub

	Public Sub SetTimeouts(lngResolveTimeout, lngConnectTimeout, lngSendTimeout, lngReceiveTimeout)
		p_XmlHttpServerReq.SetTimeouts lngResolveTimeout, lngConnectTimeout, lngSendTimeout, lngReceiveTimeout
	End Sub

	Public Function WaitForResponse() ' Optional params: [timeoutInSeconds]
		p_XmlHttpServerReq.WaitForResponse
	End Function

	Private Sub Class_Terminate()
		Set p_XmlHttpServerReq = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_HTTP_XMLHTTP_ServerRequest.vbs" Then

End If