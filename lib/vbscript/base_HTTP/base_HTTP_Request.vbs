Option Explicit

Include "base_HTTP_Constants"
Include "base_HTTP_Response"
Include "base_HTTP_Headers"
Include "base_HTTP_CookieJar"
Include "base_Sys_Logger"
Include "base_URI"

Class base_HTTP_Request
	Private p_objHttpReq, _
    		p_objHttpResp, _
    		p_objHttpHeaders, _
    		p_objCookies

	Private p_strMethod, _
		p_objUrl, _
		p_strUserAgent, _
		p_strUsername, _
		p_strPassword, _
		p_strProxyUsername, _
		p_strProxyPassword, _
		p_strProxyServer, _
		p_strProxyBypassList, _
		p_arrParams, _
		p_varData, _
		p_varFiles, _
		p_lngResolveTimeout, _
		p_lngConnectTimeout, _
		p_lngSendTimeout, _
		p_lngReceiveTimeout, _
		p_lngAsyncTimeout,_
		p_objLogger

	Private p_blnAsync, _
		p_intMaxRetries, _
		p_blnKeepAlive, _
		p_blnStoreCookies, _
		p_blnStoreResponse, _
		p_blnEncodeCookies

	Private p_blnSent, _
		p_blnRedirected

	' Public Event OnError(ByVal lngErrorNumber As Long, ByVal strErrorDescription As String)
	' Public Event OnResponseStart(ByVal lngStatus As Long, ByVal strContentType As String)
	' Public Event OnResponseDataAvailable(ByRef bytData() As Byte)
	' Public Event OnResponseFinished()

	Private Sub Class_Initialize()
		Set p_objHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
		Set p_objHttpResp = New base_HTTP_Response
		Set p_objHttpHeaders = New base_HTTP_Headers
		Set p_objCookies = New base_HTTP_CookieJar
		Set p_objUrl = New base_URI
    
		p_strMethod = ""
		p_objHttpReq.Option(WinHttpRequestOption_UserAgentString) = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.80 Safari/537.36"
		p_strUsername = ""
		p_strPassword = ""
		p_strProxyUsername = ""
		p_strProxyPassword = ""

		With p_objHttpReq
			.Option(WinHttpRequestOption_EnableHttp1_1) = True
			p_blnAsync = False
			' .SetAutoLogonPolicy AutoLogonPolicy_Never
			' .Option(WinHttpRequestOption_EnablePassportAuthentication) = False
			' .Option(WinHttpRequestOption_EnableCertificateRevocationCheck) = True
			' .Option(WinHttpRequestOption_SecureProtocols) = SecureProtocol_ALL
			' .Option(WinHttpRequestOption_SslErrorIgnoreFlags) = SslErrorFlag_Ignore_None
			.Option(WinHttpRequestOption_EnableRedirects) = False
			.Option(WinHttpRequestOption_EnableHttpsToHttpRedirects) = False
			.Option(WinHttpRequestOption_MaxAutomaticRedirects) = 10
			p_intMaxRetries = 5
			p_blnKeepAlive = False
			' .Option(WinHttpRequestOption_MaxResponseHeaderSize) = 64000
			' .Option(WinHttpRequestOption_MaxResponseDrainSize) = 1000000
			p_blnStoreCookies = True
			p_blnStoreResponse = True
			p_blnEncodeCookies = False
			' .Option(WinHttpRequestOption_URLCodePage) = UTF_8
			' .Option(WinHttpRequestOption_EscapePercentInURL) = True
			' .Option(WinHttpRequestOption_UrlEscapeDisable) = True
			' .Option(WinHttpRequestOption_UrlEscapeDisableQuery) = True
			p_lngResolveTimeout = 0
			p_lngConnectTimeout = 0
			p_lngSendTimeout = 0
			p_lngReceiveTimeout = 0
			p_lngAsyncTimeout = 0
			.SetTimeouts p_lngResolveTimeout, p_lngConnectTimeout, p_lngSendTimeout, p_lngReceiveTimeout
			Set p_objLogger = New base_Sys_Logger
			' .Option(WinHttpRequestOption_EnableTracing) = False
		End With

		p_blnSent = False
		p_blnRedirected = False
	End Sub

	
	' Properties


	Public Property Get Method()
		Method = p_strMethod
	End Property

	Public Property Let Method(strMethod)
		p_strMethod = strMethod
	End Property

	Public Property Get URL()
		URL = p_objUrl.ToString()
	End Property

	Public Property Let URL(strURL)
		Set p_objUrl = p_objUrl.FromString(strURL)
	End Property

	Public Property Set URL(objURL)
		Set p_objUrl = objURL
	End Property

	Public Property Get UserAgent()
		UserAgent = p_objHttpReq.Option(WinHttpRequestOption_UserAgentString)
	End Property

	Public Property Let UserAgent( _
		ByVal strUserAgent _
		)
    
		p_objHttpReq.Option(WinHttpRequestOption_UserAgentString) = strUserAgent
	End Property

	Public Property Get Username()
		Username = p_strUsername
	End Property

	Public Property Let Username(strUsername)
		p_strUsername = strUsername
	End Property

	Public Property Get Password()
		Password = p_strPassword
	End Property

	Public Property Let Password(strPassword)
		p_strPassword = strPassword
	End Property

	Public Property Get ProxyUsername()
		ProxyUsername = p_strProxyUsername
	End Property

	Public Property Let ProxyUsername( _
		ByVal strProxyUsername _
		)
    
		p_strProxyUsername = strProxyUsername
	End Property

	Public Property Get ProxyPassword()
		ProxyPassword = p_strProxyPassword
	End Property

	Public Property Let ProxyPassword( _
		ByVal strProxyPassword _
		)
    
		p_strProxyPassword = strProxyPassword
	End Property

	Public Property Get ProxyServer()
		ProxyServer = p_strProxyServer
	End Property

	Public Property Let ProxyServer( _
		ByVal strProxyServer _
		)
    
		p_strProxyServer = strProxyServer
	End Property

	Public Property Get ProxyBypassList()
		ProxyBypassList = p_strProxyBypassList
	End Property

	Public Property Let ProxyBypassList( _
		ByVal strProxyBypassList _
		)
    
		p_strProxyBypassList = strProxyBypassList
	End Property

	Public Property Get Data()
		Data = p_varData
	End Property

	Public Property Get Params()
		Params = p_arrParams
	End Property

	Public Property Get Files()

	End Property

	Public Property Get RawHeaders()

	End Property

	Public Property Get Headers()
		Set Headers = p_objHttpHeaders
	End Property

	Public Property Get Cookies()
		Set Cookies = p_objCookies
	End Property

	Public Property Get Response()
		Set Response = p_objHttpResp
	End Property


	' Options


	Public Property Get HttpVersion()
		If p_objHttpReq.Option(WinHttpRequestOption_EnableHttp1_1) Then
			HttpVersion = "1.1"
		Else
			HttpVersion = "1.0"
		End If
	End Property

	Public Property Let HttpVersion(strVersion)
		If strVersion = "1.1" Then
			p_objHttpReq.Option(WinHttpRequestOption_EnableHttp1_1) = True
		ElseIf strVersion = "1.0" Then
			p_objHttpReq.Option(WinHttpRequestOption_EnableHttp1_1) = False
		End If
	End Property

	Public Property Get Async()
		Async = p_blnAsync
	End Property

	Public Property Let Async(blnAsync)
		p_blnAsync = blnAsync
	End Property

	Public Property Let AutoAuth(blnAuto)
		If blnAuto Then
			p_objHttpReq.SetAutoLogonPolicy(AutoLogonPolicy_Always)
		ElseIf Not blnAuto Then
			p_objHttpReq.SetAutoLogonPolicy(AutoLogonPolicy_Never)
		End If
	End Property

	Public Property Let AutoAuthBypassProxy(blnAuto)
		If blnAuto Then
			p_objHttpReq.SetAutoLogonPolicy(AutoLogonPolicy_OnlyIfBypassProxy)
		ElseIf Not blnAuto Then
			p_objHttpReq.SetAutoLogonPolicy(AutoLogonPolicy_Never)
		End If
	End Property

	Public Property Get PassportAuth()
		PassportAuth = p_objHttpReq.Option(WinHttpRequestOption_EnablePassportAuthentication)
	End Property

	Public Property Let PassportAuth(blnPassport)
		p_objHttpReq.Option(WinHttpRequestOption_EnablePassportAuthentication) = blnPassport
	End Property

	Public Sub ClientCertification(strCert)
		p_objHttpReq.SetClientCertificate(strCert)
	End Sub

	Public Property Get ImpersonateSecureClientAuth()
		If p_objHttpReq.Option(WinHttpRequestOption_RevertImpersonationOverSsl) Then
			ImpersonateSecureClientAuth = False
		Else
			ImpersonateSecureClientAuth = True
		End If
	End Property

	Public Property Let ImpersonateSecureClientAuth(blnImpersonate)
		If blnImpersonate Then
			p_objHttpReq.Option(WinHttpRequestOption_RevertImpersonationOverSsl) = False
		Else
			p_objHttpReq.Option(WinHttpRequestOption_RevertImpersonationOverSsl) = True
		End If
	End Property

	Public Property Get VerifyServerCert()
		VerifyServerCert = p_objHttpReq.Option(WinHttpRequestOption_EnableCertificateRevocationCheck)
	End Property

	Public Property Let VerifyServerCert(blnVerify)
		p_objHttpReq.Option(WinHttpRequestOption_EnableCertificateRevocationCheck) = blnVerify
	End Property

	Public Property Let SecureProtocols(intProtocols) 
		p_objHttpReq.Option(WinHttpRequestOption_SecureProtocols) = intProtocols
	End Property

	Public Property Get Secure()
		If p_objHttpReq.Option(WinHttpRequestOption_SecureProtocols) = SecureProtocol_ALL Then
			Secure = True
		Else
			Secure = False
		End If
	End Property

	Public Property Let Secure(blnSecure)
		If blnSecure Then
			p_objHttpReq.Option(WinHttpRequestOption_SecureProtocols) = SecureProtocol_ALL
		Else
			p_objHttpReq.Option(WinHttpRequestOption_SecureProtocols) = SecureProtocol_NONE
		End If
	End Property

	Public Property Get IgnoreSSLErrors()
		IgnoreSSLErrors = p_objHttpReq.Option(WinHttpRequestOption_SslErrorIgnoreFlags)
	End Property

	Public Property Let IgnoreSSLErrors(lngIgnore)
		p_objHttpReq.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = lngIgnore
	End Property

	Public Property Get AllowRedirects()
		AllowRedirects = p_objHttpReq.Option(WinHttpRequestOption_EnableRedirects)
	End Property

	Public Property Let AllowRedirects(blnRedirect)
		p_objHttpReq.Option(WinHttpRequestOption_EnableRedirects) = blnRedirect
	End Property

	Public Property Get OnlySecureRedirects()
		OnlySecureRedirects = p_objHttpReq.Option(WinHttpRequestOption_EnableHttpsToHttpRedirects)
	End Property

	Public Property Let OnlySecureRedirects(blnSecureOnly)
		p_objHttpReq.Option(WinHttpRequestOption_EnableHttpsToHttpRedirects) = blnSecureOnly
	End Property

	Public Property Get MaxRedirects()
		MaxRedirects = p_objHttpReq.Option(WinHttpRequestOption_MaxAutomaticRedirects)
	End Property

	Public Property Let MaxRedirects(intRedirects)
		p_objHttpReq.Option(WinHttpRequestOption_MaxAutomaticRedirects) = intRedirects
	End Property

	Public Property Get MaxRetries()
		MaxRetries = p_intMaxRetries
	End Property

	Public Property Let MaxRetries(intRetries)
		p_intMaxRetries = intRetries
	End Property

	Public Property Get KeepAlive()
		KeepAlive = p_blnKeepAlive
	End Property

	Public Property Let KeepAlive(blnKeepAlive)
		p_blnKeepAlive = blnKeepAlive
	End Property

	Public Property Get MaxResponseHeader()
		MaxResponseHeader = p_objHttpReq.Option(WinHttpRequestOption_MaxResponseHeaderSize)
	End Property

	Public Property Let MaxResponseHeader(lngSize)
		p_objHttpReq.Option(WinHttpRequestOption_MaxResponseHeaderSize) = lngSize
	End Property

	Public Property Get MaxResponseBody()
		MaxResponseBody = p_objHttpReq.Option(WinHttpRequestOption_MaxResponseDrainSize)
	End Property

	Public Property Let MaxResponseBody(lngSize)
		p_objHttpReq.Option(WinHttpRequestOption_MaxResponseDrainSize) = lngSize
	End Property

	Public Property Get StoreCookies()
		StoreCookies = p_blnStoreCookies
	End Property

	Public Property Let StoreCookies(blnStoreCookies)
		p_blnStoreCookies = blnStoreCookies
	End Property

	Public Property Get StoreResponse()
		StoreResponse = p_blnStoreResponse
	End Property

	Public Property Let StoreResponse(blnStoreResp)
		p_blnStoreResponse = blnStoreResp
	End Property

	Public Property Get EncodeUrl()
		If p_objHttpReq.Option(WinHttpRequestOption_EscapePercentInURL) And _
			p_objHttpReq.Option(WinHttpRequestOption_UrlEscapeDisable) And _
			p_objHttpReq.Option(WinHttpRequestOption_UrlEscapeDisableQuery) Then
        
			EncodeURI = True
		Else
			EncodeURI = False
		End If
	End Property

	Public Property Let EncodeUrl( _
		ByVal blnEncode _
		)

		With p_objHttpReq
			.Option(WinHttpRequestOption_EscapePercentInURL) = blnEncode
			.Option(WinHttpRequestOption_UrlEscapeDisable) = blnEncode
			.Option(WinHttpRequestOption_UrlEscapeDisableQuery) = blnEncode
		End With
	End Property

	Public Property Get EncodeCookies()
		EncodeCookies = p_blnEncodeCookies
	End Property

	Public Property Let EncodeCookies(blnEncodeCookies)
		p_blnEncodeCookies = blnEncodeCookies
	End Property

	Public Property Get UrlCharacterEncoding()
		UrlCharacterEncoding = p_objHttpReq.Option(WinHttpRequestOption_URLCodePage)
	End Property

	Public Property Let UrlCharacterEncoding( _
		ByVal lngEncoding _
		)
    
		p_objHttpReq.Option(WinHttpRequestOption_URLCodePage) = lngEncoding
	End Property

	Public Property Let Timeout( _
		ByVal lngTime _
		)
    
		p_lngResolveTimeout = lngTime
		p_lngConnectTimeout = lngTime
		p_lngSendTimeout = lngTime
		p_lngReceiveTimeout = lngTime

		p_objHttpReq.SetTimeouts p_lngResolveTimeout, p_lngConnectTimeout, p_lngSendTimeout, p_lngReceiveTimeout
	End Property

	Public Property Get ResolveTimeout()
		ResolveTimeout = p_lngResolveTimeout
	End Property

	Public Property Let ResolveTimeout( _
		ByVal lngTime _
		)
    
		p_lngResolveTimeout = lngTime
		p_objHttpReq.SetTimeouts p_lngResolveTimeout, p_lngConnectTimeout, p_lngSendTimeout, p_lngReceiveTimeout
	End Property

	Public Property Get ConnectTimeout()
		ConnectTimeout = p_lngConnectTimeout
	End Property

	Public Property Let ConnectTimeout( _
		ByVal lngTime _
		)
    
		p_lngConnectTimeout = lngTime
		p_objHttpReq.SetTimeouts p_lngResolveTimeout, p_lngConnectTimeout, p_lngSendTimeout, p_lngReceiveTimeout
	End Property

	Public Property Get SendTimeout()
		SendTimeout = p_lngSendTimeout
	End Property

	Public Property Let SendTimeout( _
		ByVal lngTime _
		)
    
		p_lngSendTimeout = lngTime
		p_objHttpReq.SetTimeouts p_lngResolveTimeout, p_lngConnectTimeout, p_lngSendTimeout, p_lngReceiveTimeout
	End Property

	Public Property Get ReceiveTimeout()
		ReceiveTimeout = p_lngReceiveTimeout
	End Property

	Public Property Let ReceiveTimeout( _
		ByVal lngTime _
		)
    
		p_lngReceiveTimeout = lngTime
	End Property

	Public Property Get AsyncTimeout()
		ReceiveTimeout = p_lngAsyncTimeout
	End Property

	Public Property Let AsyncTimeout( _
		ByVal lngTime _
		)
    
		p_lngAsyncTimeout = lngTime
	End Property

	Public Property Get DefaultHeader()

	End Property 

	Public Property Let DefaultHeaders(arrHeaders)

	End Property 

	Public Property Get Logger()
		Set Logger = p_objLogger
	End Property

	Public Property Set Logger( _
		ByVal objLogger _
		)

		Set p_objLogger = objLogger
	End Property

	Public Property Get Tracing()
		Tracing = p_objHttpReq.Option(WinHttpRequestOption_EnableTracing)
	End Property

	Public Property Let Tracing(blnTrace)
		p_objHttpReq.Option(WinHttpRequestOption_EnableTracing) = blnTrace
	End Property


	' Status Flags


	Public Property Get Status()
		Status = p_objHttpReq.Status & ": " & p_objHttpReq.StatusText
	End Property

	Public Property Get StatusCode()
		StatusCode = p_objHttpReq.Status
	End Property

	Public Property Get StatusText()
		StatusText = p_objHttpReq.StatusText
	End Property

	Public Property Get Redirected()
		Redirected = p_blnRedirected
	End Property

	Public Property Get Sent()
		Sent = p_blnSent
	End Property


	' Events


	' Private Sub p_objHttpReq_OnError( _
	'	ByVal lngErrorNumber, _
	'	ByVal strErrorDescription _
	'	)
	'    
	'	RaiseEvent OnError(lngErrorNumber, strErrorDescription)
	' End Sub

	' Private Sub p_objHttpReq_OnResponseStart( _
	'	ByVal lngStatus, _
	'	ByVal strContentType _
	'	)
	'	    
	'	RaiseEvent OnResponseStart(lngStatus, strContentType)
	' End Sub

	' Private Sub p_objHttpReq_OnResponseDataAvailable( _
	'	ByRef bytData() _
	'	)
    	'
	'	RaiseEvent OnResponseDataAvailable(bytData)
	' End Sub

	' Private Sub p_objHttpReq_OnResponseFinished()
	'	RaiseEvent OnResponseFinished
	' End Sub


	' Constructors


	Public Sub Configure(objBaseHeaders, _
		objLogger, _
		blnKeepAlive, _
		intMaxRetries, _
		blnDangerMode, _
		blnSafeMode, _
		blnStrictMode, _
		blnEncodeURI, _
		blnStoreCookies _
		)

	End Sub

	Public Default Sub Prepare(strMethod, _
		strURL, _
		strUsername, _
		strPassword, _
		strProxyUser, _
		strProxyPass, _
		varParams, _
		varData, _
		varFiles, _
		objCookies, _
		objHeaders _
		)

	End Sub

	Public Sub Auth( _
		ByVal strUser, _
		ByVal strPass _
		)
    
		p_strUsername = strUser
		p_strPassword = strPass
	End Sub

	Public Sub ProxyAuth( _
		ByVal strProxyUser, _
		ByVal strProxyPath _
		)
    
		p_strProxyUser = strProxyUser
		p_strProxyPass = strProxyPath
	End Sub

	Public Sub Header()

	End Sub

	Public Sub CookieHeader(varCookies)
		Select Case TypeName(varCookies)
			Case "String":
				WScript.Echo "Added cookie header as string!"
			Case Else:
				' Error
		End Select
	End Sub

	' Public Sub Headers(arrHeaders)

	' End Sub


	' Methods


	Public Function Request( _
		ByVal strMethod, _
		ByVal strUrl, _
		ByVal varData _
		)
  
		On Error Resume Next

		Dim objFinalUrl, _
			intHeaderIndex

		p_strMethod = strMethod
		p_objUrl.FromString strUrl
		p_varData = varData	

		With p_objHttpReq
			.Open strMethod, _
				p_objUrl.ToString(), _
				p_blnAsync

			If p_objUrl.Userinfo <> "" Then
				p_strUsername = p_objUrl.Username
				p_strPassword = p_objUrl.Password
			End If

			If p_strUsername <> "" And p_strPassword <> "" Then
				.SetCredentials p_strUsername, _
						p_strPassword, _
						WinHttpRequest_SetCredentials_For_Server
			End If

			If p_strProxyUsername <> "" And p_strProxyPassword <> "" Then
				.SetCredentials p_strProxyUsername, _
						p_strProxyPassword, _
			 			WinHttpRequest_SetCredentials_For_Proxy
			 	.SetProxy WinHttpRequest_ProxySetting_Proxy, _
			 			p_strProxyServer, _
			 			p_strProxyBypassList
			End If

			If p_blnKeepAlive Then p_objHttpHeaders.AddHeaderString "Connection: Keep-Alive"

			For intHeaderIndex = 1 To p_objHttpHeaders.Count
				.SetRequestHeader p_objHttpHeaders.Item(intHeaderIndex).Name, _
							p_objHttpHeaders.Item(intHeaderIndex).Value
			Next

			If Not IsNull(p_varData) And Not IsEmpty(p_varData) Then
				.Send p_varData
			Else
				.Send
			End If

			If p_blnAsync = True Then
				If p_lngAsyncTimeout > 0 Then
					.WaitForResponse p_lngAsyncTimeout
				Else
					.WaitForResponse
				End If
			End If

			' Set objFinalUrl = New base_URI
			' objFinalUrl.FromString .Option(WinHttpRequestOption_URL)

			' If Not p_objUrl.Equals(objFinalUrl) Then p_blnRedirected = True

			' *** Add this back in: objFinalUrl.ToString(), _

			p_objHttpResp.Make .GetAllResponseHeaders(), _
			 			.ResponseBody, _
			 			.ResponseStream, _
			 			.ResponseText, _
			 			.Status, _
			 			.StatusText, _
			 			strUrl, _ 
			 			p_blnRedirected

			If p_objHttpResp.IsOk() Then p_blnSent = True
			If p_blnStoreCookies Then p_objCookies.FromResponseHeaders .GetAllResponseHeaders()
		End With

		Set Request = p_objHttpResp
		' Set objFinalUrl = Nothing

		' If Err Then PrintLn Err.Number & ": " & Err.Description & " (" & Err.Source & ")"
	End Function

	Public Function Send()
		If p_strMethod <> "" And p_objUrl.ToString() <> "" Then
			Set Send = Me.Request(p_strMethod, p_objUrl.ToString(), p_varData)
		End If
	End Function

	Public Function GetRequest( _
		ByVal strUrl _
		)

		p_strMethod = "GET"
		Set GetRequest = Me.Request(p_strMethod, strUrl, p_varData)
	End Function

	Public Function PostRequest( _
		ByVal strUrl, _
		ByVal varData _
		)

		p_strMethod = "POST"
		Set PostRequest = Me.Request(p_strMethod, strUrl, varData)
	End Function

	Public Function PutRequest( _
		ByVal strUrl, _
		ByVal varData _
		)
    
		p_strMethod = "PUT"
		Set PutRequest = Me.Request(p_strMethod, strUrl, varData)
	End Function

	Public Function HeadRequest( _
		ByVal strUrl _
		)

		p_strMethod = "HEAD"
		Set HeadRequest = Me.Request(p_strMethod, strUrl, p_varData)
	End Function

	Public Function PatchRequest( _
		ByVal strUrl, _
		ByVal varData _
		)
    
		p_strMethod = "PATCH"
		Set PatchRequest = Me.Request(p_strMethod, strUrl, varData)
	End Function

	Public Function DeleteRequest( _
		ByVal strUrl _
		)

		p_strMethod = "DELETE"
		Set DeleteRequest = Me.Request(p_strMethod, strUrl, p_varData)
	End Function

	Public Sub Download(URL, file)

	End Sub

	Public Sub Cancel()
		p_objHttpReq.Abort
	End Sub

	Public Sub ClearHeaders()

	End Sub

	Public Sub ClearDefaultHeaders()

	End Sub

	Private Sub Class_Terminate()
		Set p_objHttpReq = Nothing
		Set p_objHttpResp = Nothing
		Set p_objHttpHeaders = Nothing
		Set p_objCookies = Nothing
		Set p_objUrl = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_HTTP_Request.vbs" Then

End If