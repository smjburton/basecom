Option Explicit

' Wrapper class around the WinHttp Request object ("WinHttp.WinHttpRequest.5.1")
' The Request object. It carries out all functionality of Requests.

' See: https://msdn.microsoft.com/en-us/library/windows/desktop/aa383147(v=vs.85).aspx (HTTP authentication)
' See: https://github.com/VBA-tools/VBA-Web (for a good reference)

' Properties:

' Method
' URL
' FullURL
' PathURL
' Username
' Password
' ProxyUser
' ProxyPass
' Data
' Params
' Files
' Headers
' RawHeaders
' Cookies
' Proxy

' Options:

' HttpVersion
' UserAgent
' Async
' AutoAuth
' AutoAuthBypassProxy
' PassportAuth
' ClientCertificate
' ImpersonateSecureClientAuth
' VerifyServerCert
' SecureProtocols
' Secure
' IgnoreSSLErrors
' AllowRedirects
' OnlySecureRedirects
' MaxRedirects
' MaxRetries
' KeepAlive
' MaxResponseHeader
' MaxResponseBody
' StoreCookies
' StoreResponse
' EncodeURI
' EncodeCookies
' URLCharacterEncoding
' Timeout
' StrictMode
' SafeMode
' DangerMode
' BaseHeaders
' Logger
' Tracing
' Status
' StatusCode
' StatusText
' Redirected
' Sent

' Response object and properties:

' Response

' Methods:

' Configure
' Prepare
' Request
' Send
' Download

' GetReq
' PostReq
' PutReq
' HeadReq
' PatchReq
' DeleteReq

' Auth
' ProxyAuth
' Header (* should possibly add a AddHeader() method)
' CookieHeader (* should possible add a AddCookie() method)
' Headers (* should also add a AddHeaders() method)

' FromString
' ToString
' FromDict
' ToDict

' (* potentially add an option for multi-dimensional arrays for building an HTTP request)

' ClearHeaders
' ClearDefaultHeaders
' Timeout
' RegisterHook

Sub Include(file)
	Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")
	ExecuteGlobal FSO.OpenTextFile(file, 1).ReadAll()
	Set FSO = Nothing
End Sub

Include("base_HTTP_Constants")
Include("base_HTTP_Response")
Include("base_HTTP_Cookie")
Include("base_URI")

Class base_HTTP_Request
	Private pHttpReq, _
		pHttpResp

	' Properties
	' Properties are set by the user and do not have default values.

	Private pMethod, _
		pURL, _
		pUsername, _
		pPassword, _
		pProxyUser, _
		pProxyPass, _
		pProxy, _
		pParams, _
		pData, _
		pFiles, _
		pCookies, _
		pHeaders

	' Options
	' Options have default values unless they are overridden by the user. Setting these
	' parameters is optional.

	Private pUserAgent, _
		pAsync, _
		pMaxRetries, _
		pKeepAlive, _
		pStoreCookies, _
		pStoreResponse, _
		pEncodeURI, _
		pEncodeCookies, _
		pStoreResponse, _
		pStrictMode, _
		pSafeMode, _
		pDangerMode

	' Status Flags

	Private pSent, _
		pRedirected

	' *** Need to add a check in here to see if I can instantiate the object
	' If not, throw an error and suggest to the user to use XMLHTTP or ServerXMLHTTP
	Private Sub Class_Initialize()
		Set pHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
		Set pHttpResp = New clsHttpResponse

	' Properties
		pMethod = ""
		Set pURL = New clsURL
		pUsername = ""
		pPassword = ""
		pProxyUser = ""
		pProxyPass = ""
		Set pCookies = New clsCookieJar

	' Options
		pHttpReq.Option(WinHttpRequestOption_EnableHttp1_1) = True
		pHttpReq.Option(WinHttpRequestOption_UserAgentString) = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.80 Safari/537.36"
		pAsync = False
		pHttpReq.SetAutoLogonPolicy(AutoLogonPolicy_Never)
		pHttpReq.Option(WinHttpRequestOption_EnablePassportAuthentication) = False
		pHttpReq.Option(WinHttpRequestOption_EnableCertificateRevocationCheck) = True
		pHttpReq.Option(WinHttpRequestOption_SecureProtocols) = SecureProtocol_ALL
		pHttpReq.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = SslErrorFlag_Ignore_None
		pHttpReq.Option(WinHttpRequestOption_EnableRedirects) = False
		pHttpReq.Option(WinHttpRequestOption_EnableHttpsToHttpRedirects) = False
		pHttpReq.Option(WinHttpRequestOption_MaxAutomaticRedirects) = 10
		pMaxRetries = 5
		pKeepAlive = False
		pHttpReq.Option(WinHttpRequestOption_MaxResponseHeaderSize) = 64000
		pHttpReq.Option(WinHttpRequestOption_MaxResponseDrainSize) = 1000000
		pStoreCookies = True
		pStoreResponse = True
		pEncodeURI = False
		pEncodeCookies = False
		pHttpReq.Option(WinHttpRequestOption_URLCodePage) = 65001
		' Timeout
		pStrictMode = False
		pSafeMode = False
		pDangerMode = False
		' BaseHeaders
		' Logger
		pHttpReq.Option(WinHttpRequestOption_EnableTracing) = False

	' Status Flags
		pSent = False
		pRedirected = False
	End Sub

	
	' Properties


	' HTTP Method to use.
	Public Property Get Method()
		Method = pMethod
	End Property

	Public Property Let Method(strMethod)
		pMethod = strMethod
	End Property

	' Request URL.
	Public Property Get URL()
		URL = pURL.ToString()
	End Property

	Public Property Let URL(strURL)
		Set pURL = pURL.FromString(strURL)
	End Property

	Public Property Set URL(objURL)
		Set pURL = objURL
	End Property

	' Build the actual URL to use.
	Public Property Get FullURL()

	End Property

	' Build the path URL to use.
	Public Property Get PathURL()

	End Property

	' Username used for HTTP Basic Auth.
	Public Property Get Username()
		Username = pUsername
	End Property

	Public Property Let Username(strUsername)
		pUsername = strUsername
	End Property

	' Password used for HTTP Basic Auth.
	Public Property Get Password()
		Password = pPassword
	End Property

	Public Property Let Password(strPassword)
		pPassword = strPassword
	End Property

	' Dictionary or byte of request body data to attach to the Request.
	' Used for POST method
	Public Property Get Data()
		Data = pData
	End Property

	' Dictionary or byte of querystring data to attach to the Request.
	Public Property Get Params()

	End Property

	' for multipart encoding upload.
	' You can upload files through HTTP using the setFileUpload method. This method takes
	' a file name as the first parameter, a form name as the second parameter, and data as a
	' third optional parameter. If the third data parameter is NULL, the first file name
	' parameter is considered to be a real file on disk, and Zend\Http\Client will try to read
	' this file and upload it. If the data parameter is not NULL, the first file name parameter
	' will be sent as the file name, but no actual file needs to exist on the disk. The second
	' form name parameter is always required, and is equivalent to the “name” attribute of an
	' <input> tag, if the file was to be uploaded through an HTML form. A fourth optional
	' parameter provides the file’s content-type. If not specified, and Zend\Http\Client reads
	' the file from the disk, the mime_content_type function will be used to guess the file’s
	' content type, if it is available. In any case, the default MIME type will be application/
	' octet-stream.
	' // Uploading arbitrary data as a file
	' $text = 'this is some plain text';
	' $client->setFileUpload('some_text.txt', 'upload', $text, 'text/plain');
	' // Uploading an existing file
	' $client->setFileUpload('/tmp/Backup.tar.gz', 'bufile');
	' // Send the files
	' $client->setMethod('POST');
	' $client->send();
	' Dictionary of files to multipart upload ({filename: content}).
	Public Property Get Files()

	End Property

	' Dictionary of HTTP Headers to attach to the Request.
	' Public Property Get Headers()

	' End Property

	' Outputs the headers as a string. For example:
	' GET /docs/index.html HTTP/1.1
	' Host: www.test101.com
	' Accept: image/gif, image/jpeg, */*
	' Accept-Language: en-us
	' Accept-Encoding: gzip, deflate
	' User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)
	Public Property Get RawHeaders()

	End Property

	' CookieJar to attach to Request.
	' Allows the user to attach a cookiejar to a request
	Public Property Get Cookies()
		Cookies = pCookies
	End Property

	' Set the proxy to use for the upcoming request.
	Public Property Get Proxy()

	End Property


	' Options


	' The HTTP version to be used for the request.
	' HTTP protocol version (usually ‘1.1’ or ‘1.0’)
	Public Property Get HttpVersion()
		If pHttpReq.Option(WinHttpRequestOption_EnableHttp1_1) Then
			HttpVersion = "1.1"
		Else
			HttpVersion = "1.0"
		End If
	End Property

	Public Property Let HttpVersion(strVersion)
		If strVersion = "1.1" Then
			pHttpReq.Option(WinHttpRequestOption_EnableHttp1_1) = True
		ElseIf strVersion = "1.0" Then
			pHttpReq.Option(WinHttpRequestOption_EnableHttp1_1) = False
		End If
	End Property
	
	User agent identifier string (sent in request headers)
	Public Property Get UserAgent()
		UserAgent = pHttpReq.Option(WinHttpRequestOption_UserAgentString)
	End Property

	Public Property Let UserAgent(strUserAgent)
		pHttpReq.Option(WinHttpRequestOption_UserAgentString) = strUserAgent
	End Property

	' Determines whether the HTTP request should be sent sychronously or asynchronously.
	Public Property Get Async()
		Async = pAsync
	End Property

	Public Property Let Async(blnAsync)
		pAsync = blnAsync
	End Property

	Public Property Let AutoAuth(blnAuto)
		If blnAuto Then
			pHttpReq.SetAutoLogonPolicy(AutoLogonPolicy_Always)
		ElseIf Not blnAuto Then
			pHttpReq.SetAutoLogonPolicy(AutoLogonPolicy_Never)
		End If
	End Property

	Public Property Let AutoAuthBypassProxy(blnAuto)
		If blnAuto Then
			pHttpReq.SetAutoLogonPolicy(AutoLogonPolicy_OnlyIfBypassProxy)
		ElseIf Not blnAuto Then
			pHttpReq.SetAutoLogonPolicy(AutoLogonPolicy_Never)
		End If
	End Property

	Public Property Get PassportAuth()
		PassportAuth = pHttpReq.Option(WinHttpRequestOption_EnablePassportAuthentication)
	End Property

	Public Property Let PassportAuth(blnPassport)
		pHttpReq.Option(WinHttpRequestOption_EnablePassportAuthentication) = blnPassport
	End Property

	Public Sub ClientCertification(strCert)
		pHttpReq.SetClientCertificate(strCert)
	End Sub

	Public Property Get ImpersonateSecureClientAuth()
		If pHttpReq.Option(WinHttpRequestOption_RevertImpersonationOverSsl) Then
			ImpersonateSecureClientAuth = False
		Else
			ImpersonateSecureClientAuth = True
		End If
	End Property

	Public Property Let ImpersonateSecureClientAuth(blnImpersonate)
		If blnImpersonate Then
			pHttpReq.Option(WinHttpRequestOption_RevertImpersonationOverSsl) = False
		Else
			pHttpReq.Option(WinHttpRequestOption_RevertImpersonationOverSsl) = True
		End If
	End Property

	Public Property Get VerifyServerCert()
		VerifyServerCert = pHttpReq.Option(WinHttpRequestOption_EnableCertificateRevocationCheck)
	End Property

	Public Property Let VerifyServerCert(blnVerify)
		pHttpReq.Option(WinHttpRequestOption_EnableCertificateRevocationCheck) = blnVerify
	End Property

	Public Sub SecureProtocols(intProtocols) 
		pHttpReq.Option(WinHttpRequestOption_SecureProtocols) = intProtocols
	End Sub

	Public Property Let Secure(blnSecure)
		If blnSecure Then
			pHttpReq.Option(WinHttpRequestOption_SecureProtocols) = SecureProtocol_ALL
		Else
			pHttpReq.Option(WinHttpRequestOption_SecureProtocols) = SecureProtocol_NONE
		End If
	End Property

	Public Sub IgnoreSSLErrors(lngIgnore)
		pHttpReq.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = lngIgnore
	End Sub

	' Set to True if full redirects are allowed (e.g. re-POST-ing of data at new Location)
	Public Property Get AllowRedirects()
		AllowRedirects = pHttpReq.Option(WinHttpRequestOption_EnableRedirects)
	End Property

	Public Property Let AllowRedirects(blnRedirect)
		pHttpReq.Option(WinHttpRequestOption_EnableRedirects) = blnRedirect
	End Property

	Public Property Get OnlySecureRedirects()
		OnlySecureRedirects = pHttpReq.Option(WinHttpRequestOption_EnableHttpsToHttpRedirects)
	End Property

	Public Property Let OnlySecureRedirects(blnSecureOnly)
		pHttpReq.Option(WinHttpRequestOption_EnableHttpsToHttpRedirects) = blnSecureOnly
	End Property

	' Maximum number of redirects allowed within a request.
	Public Property Get MaxRedirects()
		MaxRedirects = pHttpReq.Option(WinHttpRequestOption_MaxAutomaticRedirects)
	End Property

	Public Property Let MaxRedirects(intRedirects)
		pHttpReq.Option(WinHttpRequestOption_MaxAutomaticRedirects) = intRedirects
	End Property

	' The number of times a request should be retried in the event of a connection failure.
	Public Property Get MaxRetries()
		MaxRetries = pMaxRetries
	End Property

	Public Property Let MaxRetries(intRetries)
		pMaxRetries = intRetries
	End Property

	' Reuse HTTP Connections through a 'Connection' header.
	Public Property Get KeepAlive()
		KeepAlive = pKeepAlive
	End Property

	Public Property Let KeepAlive(blnKeepAlive)
		pKeepAlive = blnKeepAlive
	End Property

	Public Property Get MaxResponseHeader()
		MaxResponseHeader = pHttpReq.Option(WinHttpRequestOption_MaxResponseHeaderSize)
	End Property

	Public Property Let MaxResponseHeader(lngSize)
		pHttpReq.Option(WinHttpRequestOption_MaxResponseHeaderSize) = lngSize
	End Property

	Public Property Get MaxResponseBody()
		MaxResponseBody = pHttpReq.Option(WinHttpRequestOption_MaxResponseDrainSize)
	End Property

	Public Property Let MaxResponseBody(lngSize)
		pHttpReq.Option(WinHttpRequestOption_MaxResponseDrainSize) = lngSize
	End Property

	' If StoreCookies is enabled, the request object will automatically add cookies to the jar
	' Used to manage and retain cookies between requests
	' If false, the received cookies as part of the HTTP response would be ignored.
	Public Property Get StoreCookies()
		StoreCookies = pStoreCookies
	End Property

	Public Property Let StoreCookies(blnStoreCookies)
		pStoreCookies = blnStoreCookies
	End Property

	' Whether to store last response for later retrieval with getLastResponse().
	' If set to FALSE, getLastResponse() will return NULL.
	Public Property Get StoreResponse()
		StoreResponse = pStoreResponse
	End Property

	Public Property Let StoreResponse(blnStoreResp)
		pStoreResponse = blnStoreResp
	End Property

	' If true, URIs will automatically be percent-encoded.
	' Whether to strictly adhere to RFC 3986 (in practice, this means replacing ‘+’ with ‘%20’)
	' *** Need to use WinHTTP tracing in order to verify the effect of changing these settings.
	Public Property Get EncodeURI()
		' WinHttpRequestOption_EscapePercentInURL	= True
		' WinHttpRequestOption_UrlEscapeDisable		= False
		' WinHttpRequestOption_UrlEscapeDisableQuery	= False
	End Property

	Public Property Get EncodeCookies()
		EncodeCookies = pEncodeCookies
	End Property

	Public Property Let EncodeCookies(blnEncodeCookies)
		pEncodeCookies = blnEncodeCookies
	End Property

	Public Property Get URLCharacterEncoding()
		' WinHttpRequestOption_URLCodePage
	End Property

	Public Property Let URLCharacterEncoding(strEncoding)
		' WinHttpRequestOption_URLCodePage
	End Property

	' Long integer describes the timeout of the request.
	' Connection timeout (seconds)
	' Likely need to split this into four properties.
	Public Property Get Timeout(lngTime)
		' SetTimeouts(ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout)
	End Property

	' SetAutoLogonPolicy uses Windows Authentication (formerly NTLM)
	' 	It should be set to 'never' unless the user is planning to use Windows Authentication
	' The automatic logon (auto-logon) policy determines when it is acceptable for WinHTTP to include the default credentials in a request. 
	' These default credentials are often the username and password used to log on to Microsoft Windows.
	' The auto-logon policy was implemented to prevent these credentials from being casually used to authenticate against an untrusted server. 
	' The auto-logon policy only applies to the NTLM and Negotiate authentication schemes. Credentials are never automatically transmitted with other schemes.




	' If true, Requests will do its best to follow RFCs (e.g. POST redirects).
	Public Property Get StrictMode()
		StrictMode = pStrictMode
	End Property

	Public Property Let StrictMode(blnStrictMode)
		pStrictMode = blnStrictMode
	End Property

	' If true, Requests will catch all errors.
	Public Property Get SafeMode()
		SafeMode = pSafeMode
	End Property

	Public Property Let SafeMode(blnSafeMode)
		pSafeMode = blnSafeMode
	End Property

	' If true, Requests will raise errors immediately.
	Public Property Get DangerMode()
		DangerMode = pDangerMode
	End Property

	Public Property Let DangerMode(blnDangerMode)
		pDangerMode = blnDangerMode
	End Property	

	' Default HTTP headers

	' base_headers = {"Connection":"keep-alive",
	'             "User-Agent":user_agent,
	'             "Accept-Encoding":"gzip",
 	'            "Host":"xxxxxxxxxxx",
	'             "Content-Type":"application/json; charset=UTF-8"}

	' BASE_HEADERS = {"Connection":"keep-alive",
	'                     "Accept-Encoding":"gzip",
	'                     "Host":"xxxxxxxxxxx",
	'                     "Content-Type":"application/json; charset=UTF-8"}
	
	' Stream to write request logging to.
	' Destination for streaming of received data (options: string (filename),
	' true for temp file, false/null to disable streaming)

	' Default Request Headers
	' You can set default headers that will be sent on every request:

	' Unirest\Request::defaultHeader("Header1", "Value1");
	' Unirest\Request::defaultHeader("Header2", "Value2");
	Public Sub DefaultHeader()

	End Sub

	' You can set default headers in bulk by passing an array:

	' Unirest\Request::defaultHeaders(array(
    	'	"Header1" => "Value1",
    	'	"Header2" => "Value2"
	' ));
	Public Sub DefaultHeaders(arrHeaders)

	End Sub

	Public Sub ClearDefaultHeaders()

	End Sub

	Public Property Get Logger()

	End Property

	Public Property Set Logger()

	End Property

	Public Property Get Tracing()
		Tracing = pHttpReq.Option(WinHttpRequestOption_EnableTracing)
	End Property

	Public Property Let Tracing(blnTrace)
		pHttpReq.Option(WinHttpRequestOption_EnableTracing) = blnTrace
	End Property


	' TODO:
	' Possibly add shortcut methods for common HTTP headers, such as ContentType, ExpectedType, etc.

	' SSL Verification.
	' Public Property Get Verify()
	' End Property

	' Event-handling hooks.
	' Public Property Get Hooks()
	' End Property


	' Status Flags
	' Indicate the state of the HTTP request.


	Public Property Get Status()
		Status = pHttpReq.Status & ": " & pHttpReq.StatusText
	End Property

	Public Property Get StatusCode()
		StatusCode = pHttpReq.Status
	End Property

	Public Property Get StatusText()
		StatusText = pHttpReq.StatusText
	End Property

	' True if Request is part of a redirect chain (disables history and HTTPError storage).
	Public Property Get Redirected()
		Redirected = pRedirected
	End Property

	' True if Request has been sent.
	Public Property Get Sent()
		Sent = pSent
	End Property

	
	' HTTP Response


	' Response instance, containing content and metadata of HTTP Response, once sent.
	Public Property Get Response()
		Response = pHttpResp
	End Property


	' Constructor method to configure options for an HTTP request.
	' Dictionary of configurations/options for this request.
	' Easier to pass in a dictionary of configurations than to set each flag individually
	Public Sub Configure(objBaseHeaders, _
				objLogger, _
				blnKeepAlive, _
				intMaxRetries, _
				blnDangerMode, _
				blnSafeMode, _
				blnStrictMode, _
				blnEncodeURI, _
				blnStoreCookies)

	End Sub



	' Constructor method to prepare the HTTP request.
	Public Default Sub Prepare(strMethod, _
					strURL, _
					pUsername, _
					pPassword, _
					strProxyUser, _
					strProxyPass, _
					varParams, _
					varData, _
					varFiles, _
					objCookies, _
					objHeaders)
	End Sub

	' For GET, POST, PUT, HEAD, DELETE, OPTIONS, etc.
	' *** Consider putting a select case statement in here to handle different options so that
	' variable checks are not performed on every method.
	Public Sub Request(strMethod, strURL, blnAsync, varParams, varData)
		With pHttpReq
			pMethod = strMethod

			pURL.FromString(strURL)
		
			' *** Need to check if basic auth is set in the URL.
			' // You can also specify username and password in the URI
			' $client->setUri('http://christer:secret@example.com');

			.Open strMethod, strURL, blnAsync

			If pUsername <> "" And pPassword <> "" Then
				.SetCredentials pUsername, pPassword, WinHttpRequest_SetCredentials_For_Server
			End If

			If Not IsNull(varData) Then
				.Send varData
			Else
				.Send
			End If

			If blnAsync = True Then
				' *** Need to add a check in here for the timeout as an optional parameter.
				.WaitForResponse	
			End If

			' Once a request is successfully sent, sent will equal True.
			pSent = True

			' If Not .Options(WinHttpRequestOption_URL) = strURL Then
			'	pRedirected = True
			' End If
		End With
	End Sub

	Public Sub Send()
		' *** Need to check that the method and url are set
		Request pMethod, pURL.ToString(), pAsync, pParams, pData
	End Sub

	Public Sub Download(url, file)

	End Sub

	
	' Quick access to HTTP methods


	Public Sub GetReq(strURL, strAsync, varQuery)

	End Sub

	Public Sub PostReq(strURL, strAsync, varData)

	End Sub

	Public Sub PutReq()

	End Sub

	Public Sub HeadReq()

	End Sub

	Public Sub PatchReq()

	End Sub

	Public Sub DeleteReq()

	End Sub


	' Mini Constructor Methods
	

	' Authentication tuple or object to attach to Request.
	' Equivalent in VBScript would be to pass an array
	' E.g. requests.post(url, data={}, auth=('user', 'pass'))
	' For more information, see: http://docs.python-requests.org/en/latest/user/authentication/
	Public Sub Auth(pUser, pPass)
		pUsername = pUser
		pPassword = pPass
	End Sub

	' Passing a username, password (optional), defaults to Basic Authentication
	Public Sub ProxyAuth(pProxyUser, pProxyPath)

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

	' Can set headers in bulk by passing an array:
	' Headers(array("Header1" => "Value1", "Header2" => "Value2"))
	Public Sub Headers(arrHeaders)

	End Sub

	' Clear/Reset Property Methods


	Public Sub ClearHeaders()

	End Sub

	' Likely need to split this four constructors.
	' Public Sub Timeout(lngTime)
		' SetTimeouts(ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout)
	' End Sub

	' For event hooks
	' Resources for creating callbacks in VBScript:
	' http://www.advancedqtp.com/wp-content/uploads/WLW/FunctionPointers_143CC/FunctionPointersinVBScript.pdf
	' http://www.knowledgeinbox.com/articles/vbscript/implementing-callback-in-vbscript/
	' https://msdn.microsoft.com/en-us/library/windows/desktop/aa382276(v=vs.85).aspx
	' See: http://docs.python-requests.org/en/latest/user/advanced/#event-hooks
	' Public Sub RegisterHook(event, hook)
	' End Sub

	' Could also implement these events:
	' OnError
	' OnResponseDataAvailable
	' OnResponseFinished
	' OnResponseStart
	' See: https://msdn.microsoft.com/en-us/library/ms974564

	' Request objects can either be created from the provided fromString() factory

	' $request = Request::fromString(<<<EOS
	' POST /foo HTTP/1.1
	' \r\n
	' HeaderField1: header-field-value1
	' HeaderField2: header-field-value2
	' \r\n\r\n
	' foo=bar&
	' EOS
	' );

	' $string = "GET /foo HTTP/1.1\r\n\r\nSome Content";
	' $request = Request::fromString($string);

	' $request->getMethod();    // returns Request::METHOD_GET
	' $request->getUri();       // returns Zend\Uri\Http object
	' $request->getUriString(); // returns '/foo'
	' $request->getVersion();   // returns Request::VERSION_11 or '1.1'
	' $request->getContent();   // returns 'Some Content'

	' Should also consider creating a FromDict() factory method

	' Also ToString() and ToDict()

	' Sending Raw POST Data
	' You can use a Zend\Http\Client to send raw POST data using the setRawBody() method. This method
	' takes one parameter: the data to send in the request body. When sending raw POST data, it is
	' advisable to also set the encoding type using setEncType().

	' Sending Raw POST Data

	' $xml = '<book>' .
       	'	'  <title>Islands in the Stream</title>' .
       	'	'  <author>Ernest Hemingway</author>' .
       	'	'  <year>1970</year>' .
       	'	'</book>';
	' $client->setMethod('POST');
	' $client->setRawBody($xml);
	' $client->setEncType('text/xml');
	' $client->send();

	' The data should be available on the server side through PHP‘s $HTTP_RAW_POST_DATA variable or
	' through the php://input stream.

	' Note
	' Using raw POST data
	' Setting raw POST data for a request will override any POST parameters or file uploads. You should not try to use both on the same request. Keep in mind that most servers will ignore the request body unless you send a POST request.

	' Data Streaming¶
	' By default, Zend\Http\Client accepts and returns data as PHP strings. However, in many cases there are big files to be received, thus keeping them in memory might be unnecessary or too expensive. For these cases, Zend\Http\Client supports writing data to files (streams).

	' In order to receive data from the server as stream, use setStream(). Optional argument specifies the filename where the data will be stored. If the argument is just TRUE (default), temporary file will be used and will be deleted once response object is destroyed. Setting argument to FALSE disables the streaming functionality.

	' When using streaming, send() method will return object of class Zend\Http\Response\Stream, which has two useful methods: getStreamName() will return the name of the file where the response is stored, and getStream() will return stream from which the response could be read.

	' You can either write the response to pre-defined file, or use temporary file for storing it and send it out or write it to another file using regular stream functions.

	' Receiving file from HTTP server with streaming

	' $client->setStream(); // will use temp file
	' $response = $client->send();
	' // copy file
	' copy($response->getStreamName(), "my/downloads/file");
	' // use stream
	' $fp = fopen("my/downloads/file2", "w");
	' stream_copy_to_stream($response->getStream(), $fp);
	' // Also can write to known file
	' $client->setStream("my/downloads/myfile")->send();

	Private Sub Class_Terminate()
		Set pHttpReq = Nothing
		Set pHttpResp = Nothing
		Set pCookies = Nothing
		Set pURL = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_HTTP_Request.vbs" Then
	Dim http
	Set http = New base_HTTP_Request
	
	With http
		.Auth "TheUser", "ThePass"
		.CookieHeader "Set-Cookie: sessionToken=abc123; Domain=foo.com; Path=/; Expires=Wed, 09 Jun 2021 10:18:14 GMT; Secure; HttpOnly" 
		.Request "GET", "http://www.startech.com", False, Null, Null
	End With
End If

' Example usage:
' Dim oHttp
' Set oHttp=CreateObject("WinHttp.WinHttpRequest.5.1")
' sUrl = InputBox("Url",wscript.ScriptName,"http://www.dormforce.net";)
' With oHttp
' .Open "GET",sUrl,False
' .Option(6)=0
' .SetTimeouts 5000,5000,30000,5000
' .SetRequestHeader "Accept","*/*"
' .SetRequestHeader "Accept-Language","zh-cn"
' .SetRequestHeader "User-Agent","Mozilla/4.0 (compatible; MSIE 6.0;)"
' .SetRequestHeader "HOST","www.dormforce.net"
' .SetRequestHeader "Connection","Keep-Alive"
' .Send

' Another example:
' Option Explicit
' Wscript.Echo(GetDataFromURL("http://www.808.dk/", "GET", ""))
' Function GetDataFromURL(strURL, strMethod, strPostData)
'   Dim lngTimeout
'   Dim strUserAgentString
'   Dim intSslErrorIgnoreFlags
'   Dim blnEnableRedirects
'   Dim blnEnableHttpsToHttpRedirects
'   Dim strHostOverride
'   Dim strLogin
'   Dim strPassword
'   Dim strResponseText
'   Dim objWinHttp
'   lngTimeout = 59000
'   strUserAgentString = "http_requester/0.1"
'   intSslErrorIgnoreFlags = 13056 ' 13056: ignore all err, 0: accept no err
'   blnEnableRedirects = True
'   blnEnableHttpsToHttpRedirects = True
'   strHostOverride = ""
'   strLogin = ""
'   strPassword = ""
'   Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
'   objWinHttp.SetTimeouts lngTimeout, lngTimeout, lngTimeout, lngTimeout
'   objWinHttp.Open strMethod, strURL
'   If strMethod = "POST" Then
'     objWinHttp.setRequestHeader "Content-type", _
'      "application/x-www-form-urlencoded"
'   End If
'   If strHostOverride <> "" Then
'     objWinHttp.SetRequestHeader "Host", strHostOverride
'   End If
'   objWinHttp.Option(0) = strUserAgentString
'   objWinHttp.Option(4) = intSslErrorIgnoreFlags
'   objWinHttp.Option(6) = blnEnableRedirects
'   objWinHttp.Option(12) = blnEnableHttpsToHttpRedirects
'   If (strLogin <> "") And (strPassword <> "") Then
'     objWinHttp.SetCredentials strLogin, strPassword, 0
'   End If    
'   On Error Resume Next
'   objWinHttp.Send(strPostData)
'   If Err.Number = 0 Then
'     If objWinHttp.Status = "200" Then
'       GetDataFromURL = objWinHttp.ResponseText
'     Else
'       GetDataFromURL = "HTTP " & objWinHttp.Status & " " & _
'         objWinHttp.StatusText
'     End If
'   Else
'     GetDataFromURL = "Error " & Err.Number & " " & Err.Source & " " & _
'       Err.Description
'   End If
'   On Error GoTo 0
'   Set objWinHttp = Nothing
' End Function
' And another:
' Function eBayFileExchangeUPLoad(CsvFile As String) As String
' 
'     Const URL = "https://bulksell.ebay.com/ws/eBayISAPI.dll?FileExchangeUpload"
'     Dim WinHttpReq As WinHttp.WinHttpRequest
'     Dim RequestContent As String
'     Dim MyToken As String
'     Dim ReturnString As String
'     Dim wsLine As String
' 
'     Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
'     MyToken = "****************Insert Your Token*******************"
'     
'     RequestContent = "--THIS_STRING_SEPARATES" & vbCrLf _
'     & "Content-Disposition: form-data; name=" & """token""" & vbCrLf & vbCrLf _
'     & MyToken & vbCrLf _
'     & "--THIS_STRING_SEPARATES" & vbCrLf _
'     & "Content-Disposition: form-data; name=" & """file""" & "; filename=""" & CsvFile & """" & vbCrLf _
'     & "Content-Type: text/csv" & vbCrLf & vbCrLf
'     
'     Open CsvFile For Input As #76
'     Do Until EOF(76)
'        Line Input #76, wsLine
'        RequestContent = RequestContent + wsLine + vbLf
'     Loop
'     Close #76
'        
'     RequestContent = RequestContent + vbCrLf
'     RequestContent = RequestContent & "--THIS_STRING_SEPARATES"
' 
'     MsgBox (Replace(RequestContent, MyToken, "**Token**"))
'     
'     WinHttpReq.open "POST", URL, False
'     WinHttpReq.setRequestHeader "Connection", "Keep-Alive"
'     WinHttpReq.setRequestHeader "User-Agent", "Halfupd v2"
'     WinHttpReq.setRequestHeader "Content-Type", "multipart/form-data; boundary=THIS_STRING_SEPARATES"
'     WinHttpReq.setRequestHeader "Content-Length", Len(RequestContent)
'    
'     WinHttpReq.send (RequestContent)
'      
'     ReturnString = WinHttpReq.responseText
'     eBayFileExchangeUPLoad= ReturnString
'     MsgBox (ReturnString)
' End Function

' Dim strCookie As String, strResponse As String, _
'     strUrl As String
'
'   Dim xobj As Object
'
'   Set xobj = New WinHttp.WinHttpRequest
'
'   strUrl = "https://www.example.com/login.php"
'   xobj.Open "POST", strUrl, False
'   xobj.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
'   xobj.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'   xobj.Send "username=johndoe2&password=mypassword"
'
'   strCookie = xobj.GetResponseHeader("Set-Cookie")
'   strResponse = xobj.ResponseText
'
' now try to get confidential contents:
'
'   strUrl = "https://www.example.com/secret-contents.php"
'   xobj.Open "GET", strUrl, False
'
' these 2 instructions are determining:
'
'   xobj.SetRequestHeader "Connection", "keep-alive"
'   xobj.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
'
'   xobj.SetRequestHeader "Cookie", strCookie
'   xobj.Send
'
'   strCookie = xobj.GetResponseHeader("Set-Cookie")
'   strResponse = xobj.ResponseText
' Source: http://stackoverflow.com/questions/19294956/winhttp-vba-subsequent-request-cannot-use-the-previous-login-credentials
' oHTTP.Open("GET", URL , False)
' oHTTP.SetRequestHeader("Referer", URL)
' oHTTP.Send()
'
' m_oWinHttp.SetCredentials(m_szUser, m_szPwd, WinHttp.WinHttpRequestAutoLogonPolicy.AutoLogonPolicy_Always)
' m_oWinHttp.SetClientCertificate(m_szCertificateLocation)
'
' oHttp.SetAutoLogonPolicy AutoLogonPolicy_Always
' WScript.Echo( 'User agent:      '+ WinHttpReq.Option(WinHttpRequestOption_UserAgentString));
' WScript.Echo( 'URL:             '+ WinHttpReq.Option(WinHttpRequestOption_URL));
' WScript.Echo( 'Code page:       '+ WinHttpReq.Option(WinHttpRequestOption_URLCodePage));
' WScript.Echo( 'Escape percents: '+ WinHttpReq.Option(WinHttpRequestOption_EscapePercentInURL));
