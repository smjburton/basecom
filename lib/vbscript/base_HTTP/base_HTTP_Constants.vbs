Option Explicit

' ----------------------------------------- WinHTTP Constants ----------------------------------------- '

' WinHttpRequestAutoLogonPolicy:

' SetAutoLogonPolicy uses Windows Authentication (formerly NTLM)
'   It should be set to 'never' unless the user is planning to use Windows Authentication
' The automatic logon (auto-logon) policy determines when it is acceptable for WinHTTP to include the default credentials in a request.
' These default credentials are often the username and password used to log on to Microsoft Windows.
' The auto-logon policy was implemented to prevent these credentials from being casually used to authenticate against an untrusted server.
' The auto-logon policy only applies to the NTLM and Negotiate authentication schemes. Credentials are never automatically transmitted with other schemes.

' An authenticated log on, using the default credentials, is performed for all requests.

Public Const AutoLogonPolicy_Always = 0

' Authentication is not used automatically.

Public Const AutoLogonPolicy_Never = 2

' An authenticated log on, using the default credentials, is performed only for requests on
' the local intranet. The local intranet is considered to be any server on the proxy bypass
' list in the current proxy configuration.

Public Const AutoLogonPolicy_OnlyIfBypassProxy = 1

' WinHttpRequestOption:

' Enables server certificate revocation checking during SSL negotiation. When the server presents
' a certificate, a check is performed to determine whether the certificate has been revoked by its
' issuer. If the certificate is indeed revoked, or the revocation check fails because the Certificate
' Revocation List (CRL) cannot be downloaded, the request fails; such revocation errors cannot be
' suppressed.
' Windows XP with SP1 and Windows 2000 with SP3:  This enumeration value is not supported.

Public Const WinHttpRequestOption_EnableCertificateRevocationCheck = 18

' Sets or retrieves a boolean value that indicates whether HTTP/1.1 or HTTP/1.0 should be
' used. The default is TRUE, so that HTTP/1.1 is used by default.
' Windows XP with SP1 and Windows 2000 with SP3:  This enumeration value is not supported.

Public Const WinHttpRequestOption_EnableHttp1_1 = 17

' Controls whether or not WinHTTP allows redirects. By default, all redirects are automatically
' followed, except those that transfer from a secure (https) URL to an non-secure (http) URL.
' Set this option to TRUE to enable HTTPS to HTTP redirects.

Public Const WinHttpRequestOption_EnableHttpsToHttpRedirects = 12

' Enables or disables support for Passport authentication. By default, automatic support for
' Passport authentication is disabled; set this option to TRUE to enable Passport authentication
' support.

Public Const WinHttpRequestOption_EnablePassportAuthentication = 13

' Sets or retrieves a VARIANT that indicates whether requests are automatically redirected
' when the server specifies a new location for the resource. The default value of this option
' is VARIANT_TRUE to indicate that requests are automatically redirected.

Public Const WinHttpRequestOption_EnableRedirects = 6

' Sets or retrieves a VARIANT that indicates whether tracing is currently enabled. For more
' information about the trace facility in Microsoft Windows HTTP Services (WinHTTP), see WinHTTP
' Trace Facility.

Public Const WinHttpRequestOption_EnableTracing = 10

' Sets or retrieves a VARIANT that indicates whether percent characters in the URL
' string are converted to an escape sequence. The default value of this option is
' VARIANT_TRUE which specifies all unsafe American National Standards Institute (ANSI)
' characters except the percent symbol are converted to an escape sequence.
' All unsafe characters are converted to an escape sequence including the percent symbol.
' By default, all unsafe characters except the percent symbol are converted to an escape
' sequence.

Public Const WinHttpRequestOption_EscapePercentInURL = 3

' Sets or retrieves the maximum number of redirects that WinHTTP follows; the default is
' 10. This limit prevents unauthorized sites from making the WinHTTP client stall following a
' large number of redirects.
' Windows XP with SP1 and Windows 2000 with SP3:  This enumeration value is not supported.

Public Const WinHttpRequestOption_MaxAutomaticRedirects = 14

' Sets or retrieves a bound on the amount of data that will be drained from responses in order
' to reuse a connection. The default is 1 MB.
' Windows XP with SP1 and Windows 2000 with SP3:  This enumeration value is not supported.

Public Const WinHttpRequestOption_MaxResponseDrainSize = 16

' Sets or retrieves a bound set on the maximum size of the header portion of the server's response.
' This bound protects the client from a malicious server attempting to stall the client by sending
' a response with an infinite amount of header data. The default value is 64 KB.
' Windows XP with SP1 and Windows 2000 with SP3:  This enumeration value is not supported.

Public Const WinHttpRequestOption_MaxResponseHeaderSize = 15

' Controls whether the WinHttpRequest object temporarily reverts client impersonation for the
' duration of the SSL certificate authentication operations. The default setting for the
' WinHttpRequest object is TRUE. Set this option to FALSE to keep impersonation while performing
' certificate authentication operations.

Public Const WinHttpRequestOption_RevertImpersonationOverSsl = 11

' Sets or retrieves a VARIANT that indicates which secure protocols can be used. This option selects the
' protocols acceptable to the client. The protocol is negotiated during the Secure Sockets Layer (SSL)
' handshake. This can be a combination of one or more of the following flags.

' Protocol                  Value
' SSL 2.0                   0x0008
' SSL 3.0                   0x0020
' Transport Layer Security (TLS) 1.0        0x0080
 
' The default value of this option is 0x0028, which indicates that SSL 2.0 or SSL 3.0 can be used. If
' this option is set to zero, the client and server are not able to determine an acceptable security
' protocol and the next Send results in an error.

Public Const WinHttpRequestOption_SecureProtocols = 9

' Sets a VARIANT that specifies the client certificate that is sent to a server for authentication.
' This option indicates the location, certificate store, and subject of a client certificate
' delimited with backslashes. For more information about selecting a client certificate, see SSL
' in WinHTTP.

Public Const WinHttpRequestOption_SelectCertificate = 5

' Sets or retrieves a VARIANT that indicates which server certificate errors should be ignored.
' This can be a combination of one or more of the following flags.

' Error:                                			Value:
' Unknown certification authority (CA) or untrusted root        0x0100
' Wrong usage                               			0x0200
' Invalid common name (CN)                      		0x1000
' Invalid date or certificate expired                   	0x2000
 
' The default value of this option in Version 5.1 of WinHTTP is zero, which results in no errors
' being ignored. In earlier versions of WinHTTP, the default setting was 0x3300, which resulted in
' all server certificate errors being ignored by default.

Public Const WinHttpRequestOption_SslErrorIgnoreFlags = 4

' Retrieves a VARIANT that contains the URL of the resource. This value is read-only;
' you cannot set the URL using this property. The URL cannot be read until the Open
' method is called. This option is useful for checking the URL after the Send method
' is finished to verify that any redirection occurred.

Public Const WinHttpRequestOption_URL = 1

' Sets or retrieves a VARIANT that identifies the code page for the URL string. The
' default value is the UTF-8 code page. The code page is used to convert the Unicode
' URL string, passed in the Open method, to a single-byte string representation.
' An option that can be set on the Session handle that lets the application specify
' what codepage to convert the Unicode URL string into. The default is CP_UTF8. Use
' this option if you know that the receiving server expects the URL string to use a
' particular character set that is not UTF-8.

Public Const WinHttpRequestOption_URLCodePage = 2

' Sets or retrieves a VARIANT that indicates whether unsafe characters in the path and query
' components of a URL are converted to escape sequences. The default value of this option is
' VARIANT_TRUE, which specifies that characters in the path and query are converted.
' Unsafe characters in the URL passed are not converted to escape sequences.

Public Const WinHttpRequestOption_UrlEscapeDisable = 7

' Sets or retrieves a VARIANT that indicates whether unsafe characters in the query component
' of the URL are converted to escape sequences. The default value of this option is VARIANT_TRUE,
' which specifies that characters in the query are converted.
' Unsafe characters in the query component of the URL are not converted to escape sequences.

Public Const WinHttpRequestOption_UrlEscapeDisableQuery = 8

' Sets or retrieves a VARIANT that contains the user agent string.
Public Const WinHttpRequestOption_UserAgentString = 0

' WinHttpRequestSecureProtocols:

Public Const SecureProtocol_NONE = 0
Public Const SecureProtocol_SSL2 = 8
Public Const SecureProtocol_SSL3 = 32
Public Const SecureProtocol_TLS1 = 128
Public Const SecureProtocol_ALL = 168

' WinHttpRequestSslErrorFlags:

Public Const SslErrorFlag_Ignore_None = 0
Public Const SslErrorFlag_CertCNInvalid = 4096
Public Const SslErrorFlag_CertDateInvalid = 8192
Public Const SslErrorFlag_CertWrongUsage = 512
Public Const SslErrorFlag_Ignore_All = 13056
Public Const SslErrorFlag_UnknownCA = 256

' WinHttpRequest_SetCredentials_Flags:

Public Const WinHttpRequest_SetCredentials_For_Server = 0
Public Const WinHttpRequest_SetCredentials_For_Proxy = 1

' WinHttpRequestSetProxyFlags:

Public Const WinHttpRequest_ProxySetting_Default = 0
Public Const WinHttpRequest_ProxySetting_Preconfig = 0
Public Const WinHttpRequest_ProxySetting_Direct = 1
Public Const WinHttpRequest_ProxySetting_Proxy = 2

' ----------------------------------------- HTTP Status Codes ----------------------------------------- '

' 1xx Informational:

Public Const HTTP_Continue = 100
Public Const HTTP_Switching_Protocols = 101

' 2xx Success:

Public Const HTTP_OK = 200
Public Const HTTP_Created = 201
Public Const HTTP_Accepted = 202
Public Const HTTP_NonAuthoritative_Information = 203
Public Const HTTP_No_Content = 204
Public Const HTTP_Reset_Content = 205
Public Const HTTP_Partial_Content = 206

' 3xx Redirection:

Public Const HTTP_Multiple_Choices = 300
Public Const HTTP_Moved_Permanently = 301
Public Const HTTP_Found = 302
Public Const HTTP_See_Other = 303
Public Const HTTP_Not_Modified = 304
Public Const HTTP_Use_Proxy = 305
Public Const HTTP_Switch_Proxy = 306
Public Const HTTP_Temporary_Redirect = 307

' 4xx Client Error:

Public Const HTTP_Bad_Request = 400
Public Const HTTP_Unauthorized = 401
Public Const HTTP_Payment_Required = 402
Public Const HTTP_Forbidden = 403
Public Const HTTP_Not_Found = 404
Public Const HTTP_Method_Not_Allowed = 405
Public Const HTTP_Not_Acceptable = 406
Public Const HTTP_Proxy_Authentication_Required = 407
Public Const HTTP_Request_Timeout = 408
Public Const HTTP_Conflict = 409
Public Const HTTP_Gone = 410
Public Const HTTP_Length_Required = 411
Public Const HTTP_Precondition_Failed = 412
Public Const HTTP_Payload_Too_Large = 413
Public Const HTTP_URI_Too_Long = 414
Public Const HTTP_Unsupported_Media_Type = 415
Public Const HTTP_Range_Not_Satisfiable = 416
Public Const HTTP_Expectation_Failed = 417

' 5xx Server Error:

Public Const HTTP_Internal_Server_Error = 500
Public Const HTTP_Not_Implemented = 501
Public Const HTTP_Bad_Gateway = 502
Public Const HTTP_Service_Unavailable = 503
Public Const HTTP_Gateway_Timeout = 504
Public Const HTTP_Version_Not_Supported = 505

' ----------------------------------------- URL Code Page Constants ----------------------------------------- '

Public Const IBM037 = 37                                ' IBM037  IBM EBCDIC US-Canada
Public Const IBM437 = 437                               ' OEM United States
Public Const IBM500 = 500                               ' IBM EBCDIC International
Public Const ASMO_708 = 708                             ' Arabic (ASMO 708)
Public Const ASMO_709 = 709                             ' Arabic (ASMO-449+, BCON V4)
Public Const ASMO_710 = 710                             ' Arabic - Transparent Arabic
Public Const DOS_720 = 720                              ' Arabic (Transparent ASMO); Arabic (DOS)
Public Const IBM737 = 737                               ' OEM Greek (formerly 437G); Greek (DOS)
Public Const IBM775 = 775                               ' OEM Baltic; Baltic (DOS)
Public Const IBM850 = 850                               ' OEM Multilingual Latin 1; Western European (DOS)
Public Const IBM852 = 852                               ' OEM Latin 2; Central European (DOS)
Public Const IBM855 = 855                               ' OEM Cyrillic (primarily Russian)
Public Const IBM857 = 857                               ' OEM Turkish; Turkish (DOS)
Public Const IBM00858 = 858                             ' OEM Multilingual Latin 1 + Euro symbol
Public Const IBM860 = 860                               ' OEM Portuguese; Portuguese (DOS)
Public Const IBM861 = 861                               ' OEM Icelandic; Icelandic (DOS)
Public Const DOS_862 = 862                              ' OEM Hebrew; Hebrew (DOS)
Public Const IBM863 = 863                               ' OEM French Canadian; French Canadian (DOS)
Public Const IBM864 = 864                               ' OEM Arabic; Arabic (864)
Public Const IBM865 = 865                               ' OEM Nordic; Nordic (DOS)
Public Const CP866 = 866                                ' OEM Russian; Cyrillic (DOS)
Public Const IBM869 = 869                               ' OEM Modern Greek; Greek, Modern (DOS)
Public Const IBM870 = 870                               ' IBM EBCDIC Multilingual/ROECE (Latin 2); IBM EBCDIC Multilingual Latin 2
Public Const Windows_874 = 874                          ' ANSI/OEM Thai (ISO 8859-11); Thai (Windows)
Public Const CP875 = 875                                ' IBM EBCDIC Greek Modern
Public Const Shift_JIS = 932                            ' ANSI/OEM Japanese; Japanese (Shift-JIS)
Public Const GB2312 = 936                               ' ANSI/OEM Simplified Chinese (PRC, Singapore); Chinese Simplified (GB2312)
Public Const KS_C_5601_1987 = 949                       ' ANSI/OEM Korean (Unified Hangul Code)
Public Const Big5 = 950                                 ' ANSI/OEM Traditional Chinese (Taiwan; Hong Kong SAR, PRC); Chinese Traditional (Big5)
Public Const IBM1026 = 1026                             ' IBM EBCDIC Turkish (Latin 5)
Public Const IBM01047 = 1047                            ' IBM EBCDIC Latin 1/Open System
Public Const IBM01140 = 1140                            ' IBM EBCDIC US-Canada (037 + Euro symbol); IBM EBCDIC (US-Canada-Euro)
Public Const IBM01141 = 1141                            ' IBM EBCDIC Germany (20273 + Euro symbol); IBM EBCDIC (Germany-Euro)
Public Const IBM01142 = 1142                            ' IBM EBCDIC Denmark-Norway (20277 + Euro symbol); IBM EBCDIC (Denmark-Norway-Euro)
Public Const IBM01143 = 1143                            ' IBM EBCDIC Finland-Sweden (20278 + Euro symbol); IBM EBCDIC (Finland-Sweden-Euro)
Public Const IBM01144 = 1144                            ' IBM EBCDIC Italy (20280 + Euro symbol); IBM EBCDIC (Italy-Euro)
Public Const IBM01145 = 1145                            ' IBM EBCDIC Latin America-Spain (20284 + Euro symbol); IBM EBCDIC (Spain-Euro)
Public Const IBM01146 = 1146                            ' IBM EBCDIC United Kingdom (20285 + Euro symbol); IBM EBCDIC (UK-Euro)
Public Const IBM01147 = 1147                            ' IBM EBCDIC France (20297 + Euro symbol); IBM EBCDIC (France-Euro)
Public Const IBM01148 = 1148                            ' IBM EBCDIC International (500 + Euro symbol); IBM EBCDIC (International-Euro)
Public Const IBM01149 = 1149                            ' IBM EBCDIC Icelandic (20871 + Euro symbol); IBM EBCDIC (Icelandic-Euro)
Public Const UTF_16 = 1200                              ' Unicode UTF-16, little endian byte order (BMP of ISO 10646); available only to managed applications
Public Const UnicodeFFFE = 1201                         ' Unicode UTF-16, big endian byte order; available only to managed applications
Public Const Windows_1250 = 1250                        ' ANSI Central European; Central European (Windows)
Public Const Windows_1251 = 1251                        ' ANSI Cyrillic; Cyrillic (Windows)
Public Const Windows_1252 = 1252                        ' ANSI Latin 1; Western European (Windows)
Public Const Windows_1253 = 1253                        ' ANSI Greek; Greek (Windows)
Public Const Windows_1254 = 1254                        ' ANSI Turkish; Turkish (Windows)
Public Const Windows_1255 = 1255                        ' ANSI Hebrew; Hebrew (Windows)
Public Const Windows_1256 = 1256                        ' ANSI Arabic; Arabic (Windows)
Public Const Windows_1257 = 1257                        ' ANSI Baltic; Baltic (Windows)
Public Const Windows_1258 = 1258                        ' ANSI/OEM Vietnamese; Vietnamese (Windows)
Public Const Johab = 1361                               ' Korean(Johab)
Public Const Macintosh = 10000                          ' MAC Roman; Western European (Mac)
Public Const X_Mac_Japanese = 10001                     ' Japanese (Mac)
Public Const X_Mac_ChineseTrad = 10002                  ' MAC Traditional Chinese (Big5); Chinese Traditional (Mac)
Public Const X_Mac_Korean = 10003                       ' Korean (Mac)
Public Const X_Mac_Arabic = 10004                       ' Arabic (Mac)
Public Const X_Mac_Hebrew = 10005                       ' Hebrew (Mac)
Public Const X_Mac_Greek = 10006                        ' Greek (Mac)
Public Const X_Mac_Cyrillic = 10007                     ' Cyrillic (Mac)
Public Const X_Mac_ChineseSimp = 10008                  ' MAC Simplified Chinese (GB 2312); Chinese Simplified (Mac)
Public Const X_Mac_Romanian = 10010                     ' Romanian (Mac)
Public Const X_Mac_Ukrainian = 10017                    ' Ukrainian (Mac)
Public Const X_Mac_Thai = 10021                         ' Thai (Mac)
Public Const X_Mac_CE = 10029                           ' MAC Latin 2; Central European (Mac)
Public Const X_Mac_Icelandic = 10079                    ' Icelandic (Mac)
Public Const X_Mac_Turkish = 10081                      ' Turkish (Mac)
Public Const X_Mac_Croatian = 10082                     ' Croatian (Mac)
Public Const UTF_32 = 12000                             ' Unicode UTF-32, little endian byte order; available only to managed applications
Public Const UTF_32BE = 12001                           ' Unicode UTF-32, big endian byte order; available only to managed applications
Public Const X_Chinese_CNS = 20000                      ' CNS Taiwan; Chinese Traditional (CNS)
Public Const X_CP20001 = 20001                          ' TCA Taiwan
Public Const X_Chinese_Eten = 20002                     ' Eten Taiwan; Chinese Traditional (Eten)
Public Const X_CP20003 = 20003                          ' IBM5550 Taiwan
Public Const X_CP20004 = 20004                          ' TeleText Taiwan
Public Const X_CP20005 = 20005                          ' Wang Taiwan
Public Const X_IA5 = 20105                              ' IA5 (IRV International Alphabet No. 5, 7-bit); Western European (IA5)
Public Const X_IA5_German = 20106                       ' IA5 German (7-bit)
Public Const X_IA5_Swedish = 20107                      ' IA5 Swedish (7-bit)
Public Const X_IA5_Norwegian = 20108                    ' IA5 Norwegian (7-bit)
Public Const US_ASCII = 20127                           ' US-ASCII (7-bit)
Public Const X_CP20261 = 20261                          ' T.61
Public Const X_CP20269 = 20269                          ' ISO 6937 Non-Spacing Accent
Public Const IBM273 = 20273                             ' IBM EBCDIC Germany
Public Const IBM277 = 20277                             ' IBM EBCDIC Denmark-Norway
Public Const IBM278 = 20278                             ' IBM EBCDIC Finland-Sweden
Public Const IBM280 = 20280                             ' IBM EBCDIC Italy
Public Const IBM284 = 20284                             ' IBM EBCDIC Latin America-Spain
Public Const IBM285 = 20285                             ' IBM EBCDIC United Kingdom
Public Const IBM290 = 20290                             ' IBM EBCDIC Japanese Katakana Extended
Public Const IBM297 = 20297                             ' IBM EBCDIC France
Public Const IBM420 = 20420                             ' IBM EBCDIC Arabic
Public Const IBM423 = 20423                             ' IBM EBCDIC Greek
Public Const IBM424 = 20424                             ' IBM EBCDIC Hebrew
Public Const X_EBCDIC_KoreanExtended = 20833            ' IBM EBCDIC Korean Extended
Public Const IBM_Thai = 20838                           ' IBM EBCDIC Thai
Public Const KOI8_R = 20866                             ' Russian (KOI8-R); Cyrillic (KOI8-R)
Public Const IBM871 = 20871                             ' IBM EBCDIC Icelandic
Public Const IBM880 = 20880                             ' IBM EBCDIC Cyrillic Russian
Public Const IBM905 = 20905                             ' IBM EBCDIC Turkish
Public Const IBM00924 = 20924                           ' IBM EBCDIC Latin 1/Open System (1047 + Euro symbol)
Public Const EUC_JP_JIS = 20932                         ' Japanese (JIS 0208-1990 and 0212-1990)
Public Const X_CP20936 = 20936                          ' Simplified Chinese (GB2312); Chinese Simplified (GB2312-80)
Public Const X_CP20949 = 20949                          ' Korean Wansung
Public Const CP1025 = 21025                             ' IBM EBCDIC Cyrillic Serbian-Bulgarian
Public Const KOI8_U = 21866                             ' Ukrainian (KOI8-U); Cyrillic (KOI8-U)
Public Const ISO_8859_1 = 28591                         ' ISO 8859-1 Latin 1; Western European (ISO)
Public Const ISO_8859_2 = 28592                         ' ISO 8859-2 Central European; Central European (ISO)
Public Const ISO_8859_3 = 28593                         ' ISO 8859-3 Latin 3
Public Const ISO_8859_4 = 28594                         ' ISO 8859-4 Baltic
Public Const ISO_8859_5 = 28595                         ' ISO 8859-5 Cyrillic
Public Const ISO_8859_6 = 28596                         ' ISO 8859-6 Arabic
Public Const ISO_8859_7 = 28597                         ' ISO 8859-7 Greek
Public Const ISO_8859_8 = 28598                         ' ISO 8859-8 Hebrew; Hebrew (ISO-Visual)
Public Const ISO_8859_9 = 28599                         ' ISO 8859-9 Turkish
Public Const ISO_8859_13 = 28603                        ' ISO 8859-13 Estonian
Public Const ISO_8859_15 = 28605                        ' ISO 8859-15 Latin 9
Public Const X_Europa = 29001                           ' Europa 3
Public Const ISO_8859_8_i = 38598                       ' ISO 8859-8 Hebrew; Hebrew (ISO-Logical)
Public Const ISO_2022_JP_JIS = 50220                    ' ISO 2022 Japanese with no halfwidth Katakana; Japanese (JIS)
Public Const CSISO2022JP = 50221                        ' ISO 2022 Japanese with halfwidth Katakana; Japanese (JIS-Allow 1 byte Kana)
Public Const ISO_2022_JP_JISX = 50222                   ' ISO 2022 Japanese JIS X 0201-1989; Japanese (JIS-Allow 1 byte Kana - SO/SI)
Public Const ISO_2022_KR = 50225                        ' ISO 2022 Korean
Public Const X_CP50227 = 50227                          ' ISO 2022 Simplified Chinese; Chinese Simplified (ISO 2022)
Public Const X_CP50229 = 50229                          ' ISO 2022 Traditional Chinese
Public Const EBCDIC_JP = 50930                          ' EBCDIC Japanese (Katakana) Extended
Public Const EBCDIC_USCAJP = 50931                      ' EBCDIC US - Canada And Japanese
Public Const EBCDIC_KR = 50933                          ' EBCDIC Korean Extended and Korean
Public Const EBCDIC_XCP = 50935                         ' EBCDIC Simplified Chinese Extended and Simplified Chinese
Public Const EBCDIC_CN = 50936                          ' EBCDIC Simplified Chinese
Public Const EBCDIC_USCACN = 50937                      ' EBCDIC US-Canada and Traditional Chinese
Public Const EBCDIC_XJP = 50939                         ' EBCDIC Japanese (Latin) Extended and Japanese
Public Const EUC_JP = 51932                             ' EUC Japanese
Public Const EUC_CN = 51936                             ' EUC Simplified Chinese; Chinese Simplified (EUC)
Public Const EUC_KR = 51949                             ' EUC Korean
Public Const EUC = 51950                                ' Traditional Chinese
Public Const HZ_GB_2312 = 52936                         ' HZ-GB2312 Simplified Chinese; Chinese Simplified (HZ)
Public Const GB18030 = 54936                            ' Windows XP and later: GB18030 Simplified Chinese (4 byte); Chinese Simplified (GB18030)
Public Const X_ISCII_DE = 57002                         ' ISCII Devanagari
Public Const X_ISCII_BE = 57003                         ' ISCII Bangla
Public Const X_ISCII_TA = 57004                         ' ISCII Tamil
Public Const X_ISCII_TE = 57005                         ' ISCII Telugu
Public Const X_ISCII_AS = 57006                         ' ISCII Assamese
Public Const X_ISCII_OR = 57007                         ' ISCII Odia
Public Const X_ISCII_KA = 57008                         ' ISCII Kannada
Public Const X_ISCII_MA = 57009                         ' ISCII Malayalam
Public Const X_ISCII_GU = 57010                         ' ISCII Gujarati
Public Const X_ISCII_PA = 57011                         ' ISCII Punjabi
Public Const UTF_7 = 65000                              ' Unicode (UTF-7)
Public Const UTF_8 = 65001                              ' Unicode (UTF-8)