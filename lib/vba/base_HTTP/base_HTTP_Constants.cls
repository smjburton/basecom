Option Explicit

' ----------------------------------------- WinHTTP Constants -----------------------------------------

' WinHttpRequestAutoLogonPolicy:

' SetAutoLogonPolicy uses Windows Authentication (formerly NTLM)
' 	It should be set to 'never' unless the user is planning to use Windows Authentication
' The automatic logon (auto-logon) policy determines when it is acceptable for WinHTTP to include the default credentials in a request. 
' These default credentials are often the username and password used to log on to Microsoft Windows.
' The auto-logon policy was implemented to prevent these credentials from being casually used to authenticate against an untrusted server. 
' The auto-logon policy only applies to the NTLM and Negotiate authentication schemes. Credentials are never automatically transmitted with other schemes.

' An authenticated log on, using the default credentials, is performed for all requests.

Const AutoLogonPolicy_Always 						= 0

' Authentication is not used automatically.

Const AutoLogonPolicy_Never 						= 2

' An authenticated log on, using the default credentials, is performed only for requests on
' the local intranet. The local intranet is considered to be any server on the proxy bypass
' list in the current proxy configuration.

Const AutoLogonPolicy_OnlyIfBypassProxy 				= 1

' WinHttpRequestOption:

' Enables server certificate revocation checking during SSL negotiation. When the server presents
' a certificate, a check is performed to determine whether the certificate has been revoked by its
' issuer. If the certificate is indeed revoked, or the revocation check fails because the Certificate
' Revocation List (CRL) cannot be downloaded, the request fails; such revocation errors cannot be
' suppressed.
' Windows XP with SP1 and Windows 2000 with SP3:  This enumeration value is not supported.

Const WinHttpRequestOption_EnableCertificateRevocationCheck 		= 18

' Sets or retrieves a boolean value that indicates whether HTTP/1.1 or HTTP/1.0 should be
' used. The default is TRUE, so that HTTP/1.1 is used by default.
' Windows XP with SP1 and Windows 2000 with SP3:  This enumeration value is not supported.

Const WinHttpRequestOption_EnableHttp1_1 				= 17

' Controls whether or not WinHTTP allows redirects. By default, all redirects are automatically
' followed, except those that transfer from a secure (https) URL to an non-secure (http) URL.
' Set this option to TRUE to enable HTTPS to HTTP redirects.

Const WinHttpRequestOption_EnableHttpsToHttpRedirects 			= 12

' Enables or disables support for Passport authentication. By default, automatic support for
' Passport authentication is disabled; set this option to TRUE to enable Passport authentication
' support.

Const WinHttpRequestOption_EnablePassportAuthentication 		= 13

' Sets or retrieves a VARIANT that indicates whether requests are automatically redirected
' when the server specifies a new location for the resource. The default value of this option
' is VARIANT_TRUE to indicate that requests are automatically redirected.

Const WinHttpRequestOption_EnableRedirects 				= 6

' Sets or retrieves a VARIANT that indicates whether tracing is currently enabled. For more
' information about the trace facility in Microsoft Windows HTTP Services (WinHTTP), see WinHTTP
' Trace Facility.

Const WinHttpRequestOption_EnableTracing 				= 10

' Sets or retrieves a VARIANT that indicates whether percent characters in the URL
' string are converted to an escape sequence. The default value of this option is
' VARIANT_TRUE which specifies all unsafe American National Standards Institute (ANSI)
' characters except the percent symbol are converted to an escape sequence.
' All unsafe characters are converted to an escape sequence including the percent symbol.
' By default, all unsafe characters except the percent symbol are converted to an escape
' sequence.

Const WinHttpRequestOption_EscapePercentInURL 				= 3

' Sets or retrieves the maximum number of redirects that WinHTTP follows; the default is
' 10. This limit prevents unauthorized sites from making the WinHTTP client stall following a
' large number of redirects.
' Windows XP with SP1 and Windows 2000 with SP3:  This enumeration value is not supported.

Const WinHttpRequestOption_MaxAutomaticRedirects 			= 14

' Sets or retrieves a bound on the amount of data that will be drained from responses in order
' to reuse a connection. The default is 1 MB.
' Windows XP with SP1 and Windows 2000 with SP3:  This enumeration value is not supported.

Const WinHttpRequestOption_MaxResponseDrainSize 			= 16

' Sets or retrieves a bound set on the maximum size of the header portion of the server's response.
' This bound protects the client from a malicious server attempting to stall the client by sending
' a response with an infinite amount of header data. The default value is 64 KB.
' Windows XP with SP1 and Windows 2000 with SP3:  This enumeration value is not supported.

Const WinHttpRequestOption_MaxResponseHeaderSize 			= 15

' Controls whether the WinHttpRequest object temporarily reverts client impersonation for the
' duration of the SSL certificate authentication operations. The default setting for the
' WinHttpRequest object is TRUE. Set this option to FALSE to keep impersonation while performing
' certificate authentication operations.

Const WinHttpRequestOption_RevertImpersonationOverSsl 			= 11

' Sets or retrieves a VARIANT that indicates which secure protocols can be used. This option selects the
' protocols acceptable to the client. The protocol is negotiated during the Secure Sockets Layer (SSL)
' handshake. This can be a combination of one or more of the following flags.

' Protocol					Value
' SSL 2.0					0x0008
' SSL 3.0					0x0020
' Transport Layer Security (TLS) 1.0		0x0080
 
' The default value of this option is 0x0028, which indicates that SSL 2.0 or SSL 3.0 can be used. If
' this option is set to zero, the client and server are not able to determine an acceptable security
' protocol and the next Send results in an error.

Const WinHttpRequestOption_SecureProtocols 				= 9

' Sets a VARIANT that specifies the client certificate that is sent to a server for authentication.
' This option indicates the location, certificate store, and subject of a client certificate
' delimited with backslashes. For more information about selecting a client certificate, see SSL
' in WinHTTP.

Const WinHttpRequestOption_SelectCertificate 				= 5

' Sets or retrieves a VARIANT that indicates which server certificate errors should be ignored.
' This can be a combination of one or more of the following flags.

' Error:								Value:
' Unknown certification authority (CA) or untrusted root		0x0100
' Wrong usage								0x0200
' Invalid common name (CN)						0x1000
' Invalid date or certificate expired					0x2000
 
' The default value of this option in Version 5.1 of WinHTTP is zero, which results in no errors
' being ignored. In earlier versions of WinHTTP, the default setting was 0x3300, which resulted in
' all server certificate errors being ignored by default.

Const WinHttpRequestOption_SslErrorIgnoreFlags 				= 4

' Retrieves a VARIANT that contains the URL of the resource. This value is read-only;
' you cannot set the URL using this property. The URL cannot be read until the Open
' method is called. This option is useful for checking the URL after the Send method
' is finished to verify that any redirection occurred.

Const WinHttpRequestOption_URL 						= 1

' Sets or retrieves a VARIANT that identifies the code page for the URL string. The
' default value is the UTF-8 code page. The code page is used to convert the Unicode
' URL string, passed in the Open method, to a single-byte string representation.
' An option that can be set on the Session handle that lets the application specify
' what codepage to convert the Unicode URL string into. The default is CP_UTF8. Use
' this option if you know that the receiving server expects the URL string to use a
' particular character set that is not UTF-8.

Const WinHttpRequestOption_URLCodePage					= 2

' Sets or retrieves a VARIANT that indicates whether unsafe characters in the path and query
' components of a URL are converted to escape sequences. The default value of this option is
' VARIANT_TRUE, which specifies that characters in the path and query are converted.
' Unsafe characters in the URL passed are not converted to escape sequences.

Const WinHttpRequestOption_UrlEscapeDisable 				= 7

' Sets or retrieves a VARIANT that indicates whether unsafe characters in the query component
' of the URL are converted to escape sequences. The default value of this option is VARIANT_TRUE,
' which specifies that characters in the query are converted.
' Unsafe characters in the query component of the URL are not converted to escape sequences.

Const WinHttpRequestOption_UrlEscapeDisableQuery 			= 8

' Sets or retrieves a VARIANT that contains the user agent string.
Const WinHttpRequestOption_UserAgentString 				= 0

' WinHttpRequestSecureProtocols:

Const SecureProtocol_NONE 						= 0
Const SecureProtocol_SSL2 						= 8
Const SecureProtocol_SSL3 						= 32
Const SecureProtocol_TLS1 						= 128
Const SecureProtocol_ALL 						= 168

' WinHttpRequestSslErrorFlags:

Const SslErrorFlag_Ignore_None 						= 0
Const SslErrorFlag_CertCNInvalid 					= 4096
Const SslErrorFlag_CertDateInvalid 					= 8192
Const SslErrorFlag_CertWrongUsage 					= 512
Const SslErrorFlag_Ignore_All 						= 13056
Const SslErrorFlag_UnknownCA 						= 256

' WinHttpRequest_SetCredentials_Flags:

Const WinHttpRequest_SetCredentials_For_Server 				= 0
Const WinHttpRequest_SetCredentials_For_Proxy 				= 1

' ----------------------------------------- HTTP Status Codes -----------------------------------------

' 1xx Informational:

Const HTTP_Continue				= 100
Const HTTP_Switching_Protocols			= 101

' 2xx Success:

Const HTTP_OK					= 200
Const HTTP_Created				= 201
Const HTTP_Accepted				= 202
Const HTTP_NonAuthoritative_Information		= 203
Const HTTP_No_Content				= 204
Const HTTP_Reset_Content			= 205
Const HTTP_Partial_Content			= 206

' 3xx Redirection:

Const HTTP_Multiple_Choices			= 300
Const HTTP_Moved_Permanently			= 301
Const HTTP_Found				= 302
Const HTTP_See_Other				= 303
Const HTTP_Not_Modified				= 304
Const HTTP_Use_Proxy				= 305
Const HTTP_Switch_Proxy				= 306
Const HTTP_Temporary_Redirect			= 307

' 4xx Client Error:

Const HTTP_Bad_Request				= 400	
Const HTTP_Unauthorized				= 401
Const HTTP_Payment_Required			= 402
Const HTTP_Forbidden				= 403
Const HTTP_Not_Found				= 404
Const HTTP_Method_Not_Allowed			= 405
Const HTTP_Not_Acceptable			= 406 
Const HTTP_Proxy_Authentication_Required	= 407
Const HTTP_Request_Timeout			= 408
Const HTTP_Conflict				= 409
Const HTTP_Gone					= 410
Const HTTP_Length_Required			= 411
Const HTTP_Precondition_Failed			= 412
Const HTTP_Payload_Too_Large			= 413
Const HTTP_URI_Too_Long				= 414
Const HTTP_Unsupported_Media_Type		= 415
Const HTTP_Range_Not_Satisfiable		= 416
Const HTTP_Expectation_Failed			= 417

' 5xx Server Error:

Const HTTP_Internal_Server_Error		= 500
Const HTTP_Not_Implemented			= 501
Const HTTP_Bad_Gateway				= 502
Const HTTP_Service_Unavailable			= 503
Const HTTP_Gateway_Timeout			= 504
Const HTTP_Version_Not_Supported		= 505

' ---------------------------------------------------------------------------------------------------