Option Explicit

Include "base_HTTP_Headers"
Include "base_HTTP_CookieJar"
Include "base_HTTP_URI"
Include "base_JSON"

Class base_HTTP_Response
	Private p_objHttpHeaders, _
		p_objCookies, _
		p_objUrl
        
	Private p_strVersion, _
		p_varResponseBody, _
		p_varResponseStream, _
		p_strResponseText, _
		p_lngStatusCode, _
		p_strStatusReason, _
		p_blnRedirected


	' Constructors


	Private Sub Class_Initialize()
		Set p_objHttpHeaders = New base_HTTP_Headers
		Set p_objCookies = New base_HTTP_CookieJar
		Set p_objUrl = New base_URI
	End Sub

	Public Sub Make( _
		ByVal strResponseHeaders, _
		ByVal varResponseBody, _
		ByVal varResponseStream, _
		ByVal strResponseText, _
		ByVal lngStatusCode, _
		ByVal strStatusReason, _
		ByVal strUrl, _
		ByVal blnRedirected _
		)

		p_objHttpHeaders.FromString strResponseHeaders
		p_objCookies.FromResponseHeaders strResponseHeaders
		p_varResponseBody = varResponseBody
		p_varResponseStream = varResponseStream
		p_strResponseText = strResponseText
		p_lngStatusCode = lngStatusCode
		p_strStatusReason = strStatusReason
		p_objUrl.FromString strUrl
		p_blnRedirected = blnRedirected
	End Sub


	' Properties


	Public Property Get Body()
		If TypeName(p_varResponseBody) = "Object" Then
			Set Body = p_varResponseBody
		Else
			Body = p_varResponseBody
		End If
	End Property

	Public Property Get Config()

	End Property

	Public Property Get Content()

	End Property

	Public Property Get Cookies()
		Set Cookies = p_objCookies
	End Property

	Public Property Get Encoding()

	End Property

	Public Property Get Error()

	End Property

	Public Property Get Headers()
		Set Headers = p_objHttpHeaders
	End Property

	Public Property Get History()

	End Property

	Public Property Get HTML()
		' Need to check that the response type is HTML and then return an object
	End Property

	Public Property Get IsRedirect()
		IsRedirect = p_blnRedirected
	End Property

	Public Property Get IsOk()
		If p_lngStatusCode >= 200 And p_lngStatusCode < 400 Then
			IsOk = True
		Else
			IsOk = False
		End If
	End Property

	Public Property Get IsSuccess()
		If p_lngStatusCode >= 200 And p_lngStatusCode < 300 Then
			IsSuccess = True
		Else
			IsSuccess = False
		End If
	End Property

	Public Property Get IsInformational()
		If p_lngStatusCode >= 100 And p_lngStatusCode < 200 Then
			IsSuccess = True
		Else
			IsSuccess = False
		End If
	End Property

	Public Property Get IsClientError()
		If p_lngStatusCode >= 400 And p_lngStatusCode < 500 Then
			IsSuccess = True
		Else
			IsSuccess = False
		End If
	End Property

	Public Property Get IsServerError()
		If p_lngStatusCode >= 500 And p_lngStatusCode < 600 Then
			IsSuccess = True
		Else
			IsSuccess = False
		End If
	End Property

	Public Property Get IsNotFound()
		If p_lngStatusCode = HTTP_Not_Found Then
			IsSuccess = True
		Else
			IsSuccess = False
		End If
	End Property

	Public Property Get IsForbidden()
		If p_lngStatusCode = HTTP_Forbidden Then
			IsSuccess = True
		Else
			IsSuccess = False
		End If
	End Property

	Public Property Get JSON()
		' Need to check that the response type is JSON and then return an object
	End Property

	Public Property Get Raw()

	End Property

	Public Property Get Request()

	End Property

	Public Property Get Status()

	End Property

	Public Property Get StatusCode()
		StatusCode = p_lngStatusCode
	End Property

	Public Property Get StatusReason()
		StatusReason = p_strStatusReason
	End Property

	Public Property Get Stream()
		If TypeName(p_varResponseStream) = "Object" Then
			Set Stream = p_varResponseStream
		Else
			Stream = p_varResponseStream
		End If
	End Property

	Public Property Get Text()
		Text = p_strResponseText
	End Property

	Public Property Get URL()
    		Set URL = p_objUrl
	End Property

	Public Property Get XML()
		' Need to check that the response type is XML and then return an object
	End Property

	Private Sub Class_Terminate()
		Set p_objHttpHeaders = Nothing
		Set p_objCookies = Nothing
		Set p_objUrl = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_HTTP_Response.vbs" Then

End If
