Option Explicit

' “version-independent” ProgID for MSXML:
' Microsoft.XMLHTTP = IXMLHTTPRequest object
' Microsoft.XMLHTTP.1.0 = IXMLHTTPRequest object

' “version-independent” ProgID for MSXML:
' Msxml2.XMLHTTP = IXMLHTTPRequest object

' Msxml2.XMLHTTP.3.0 = IXMLHTTPRequest object
' Msxml2.XMLHTTP.4.0 = (?) - do not use; contains issues
' Msxml2.XMLHTTP.6.0 = IServerXMLHTTPRequest2 object

Class base_HTTP_XMLHTTP_Request
	Private p_XmlHttpReq

	Private Sub Class_Initialize()
		Set p_XmlHttpReq = CreateObject("MSXML2.XMLHTTP")
	End Sub

	' Properties:

	Public Property Get ReadyState()
		ReadyState = p_XmlHttpReq.ReadyState
	End Property

	Public Property Get ResponseBody()
		If IsObject(p_XmlHttpReq.ResponseBody) Then
			Set ResponseBody = p_XmlHttpReq.ResponseBody
		Else
			ResponseBody = p_XmlHttpReq.ResponseBody
		End If
	End Property

	Public Property Get ResponseStream()
		If IsObject(p_XmlHttpReq.ResponseStream) Then
			Set ResponseStream = p_XmlHttpReq.ResponseStream
		Else
			ResponseStream = p_XmlHttpReq.ResponseStream
		End If
	End Property

	Public Property Get ResponseText()
		ResponseText = p_XmlHttpReq.ResponseText
	End Property

	Public Property Get ResponseXML()
		Set ResponseXML = p_XmlHttpReq.ResponseXML
	End Property

	Public Property Get Status()
		Status = p_XmlHttpReq.Status
	End Property

	Public Property Get StatusText()
		StatusText = p_XmlHttpReq.StatusText
	End Property


	' Methods


	Public Sub Abort()
		p_XmlHttpReq.Abort
	End Sub

	Public Function GetAllResponseHeaders()
		GetAllResponseHeaders = p_XmlHttpReq.GetAllResponseHeaders()
	End Function

	Public Function GetResponseHeader(strHeader)
		GetResponseHeader = p_XmlHttpReq.GetResponseHeader()
	End Function

	Public Sub Open(strMethod, strUrl) ' Optional params: [varAsync], [bstrUser], [bstrPassword]
		p_XmlHttpReq.Open strMethod, strUrl
	End Sub

	Public Sub Send() ' Optional params: [varBody]
		p_XmlHttpReq.Send
	End Sub

	Public Sub SetRequestHeader(strHeader, strValue)
		p_XmlHttpReq.SetRequestHeader strHeader, strValue
	End Sub

	Private Sub Class_Terminate()
		Set p_XmlHttpReq = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_HTTP_XMLHTTP_Request.vbs" Then

End If