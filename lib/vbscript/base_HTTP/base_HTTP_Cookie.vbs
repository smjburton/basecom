Option Explicit

' A class representing an HTTP cookie.
' This class provides methods for parsing HTTP response strings and easily accessing their properties.
' It also allows a user to check if a cookie matches against a specific scenario such as a request URL, 
' expiration time, secure connection, etc.
' The cookie object only stores one key-value pair per object (a cookie's definition)

' *** Some cookie objects include methods for encoding and decoding cookie strings - need to find out what
' these are used for. I believe this is for URL encoding/decoding...

' See: https://docs.python.org/2/library/cookielib.html#cookielib.Cookie
' See: https://en.wikipedia.org/wiki/HTTP_cookie
' See: https://tools.ietf.org/html/rfc6265 (latest RFC spec on cookies)
' See: http://curl.haxx.se/rfc/cookie_spec.html

' For domain matching cookies, see:
' https://en.wikipedia.org/wiki/HTTP_cookie#Domain_and_Path
' http://stackoverflow.com/questions/1062963/how-do-browser-cookie-domains-work
' http://tools.ietf.org/html/rfc6265#section-5.1.3
' http://www.sitepoint.com/domain-www-or-no-www/
' http://webmasters.stackexchange.com/questions/55790/setting-cookies-only-on-the-naked-domain
' http://stackoverflow.com/questions/18492576/share-cookie-between-subdomain-and-domain
' http://erik.io/blog/2014/03/04/definitive-guide-to-cookie-domains/
' http://serverfault.com/questions/153409/can-subdomain-example-com-set-a-cookie-that-can-be-read-by-example-com

Sub Include(file)
	Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")
	ExecuteGlobal FSO.OpenTextFile(file, 1).ReadAll()
	Set FSO = Nothing
End Sub

Include "url.vbs"

Class v_HTTP_Cookie
	Private pName, _
		pValue, _
		pDomain, _
		pPath, _
		pExpires, _
		pComment, _
		pSecure, _
		pHttpOnly, _
		pVersion

	Private pDomainRegex

	Private Sub Class_Initialize()
		pName = ""
		pValue = ""
		pDomain = ""
		pPath = ""
		pExpires = 0
		pComment = ""
		pSecure = False
		pHttpOnly = False
		pVersion = 1

		Set pDomainRegex = New RegExp
	End Sub

	' The name of the cookie.
	Public Property Get Name()
		Name = pName
	End Property

	' The value of the cookie.
	Public Property Get Value()
		Value = pValue
	End Property

	' The domain that the cookie is available to. Setting the domain to 'www.example.com'
	' will make the cookie available in the www subdomain and higher subdomains. Cookies
	' available to a lower domain, such as 'example.com' will be available to higher subdomains,
	' such as 'www.example.com'. 
	Public Property Get Domain()
		Domain = pDomain
	End Property

	' The path on the server in which the cookie will be available on. If set to '/', the
	' cookie will be available within the entire domain. If set to '/foo/', the cookie will
	' only be available within the /foo/ directory and all sub-directories such as /foo/bar/
	' of domain. The default value is the current directory that the cookie is being set in.
	Public Property Get Path()
		Path = pPath
	End Property

	' The time the cookie expires. 
	' Defaults to 0, meaning "until the browser is closed".
	' It should be a datetime object. When rendering, it should always be produced in
	' the standard format.
	Public Property Get Expires()
		Expires = pExpires
	End Property

	' String comment from the server explaining the function of this cookie, or None.
	Public Property Get Comment()
		Comment = pComment
	End Property

	' Indicates that the cookie should only be transmitted over a secure HTTPS connection
	' from the client. When set to TRUE, the cookie will only be set if a secure connection
	' exists.
	Public Property Get Secure()
		Secure = pSecure
	End Property

	' When TRUE the cookie will be made accessible only through the HTTP protocol. This
	' means that the cookie won't be accessible by scripting languages, such as JavaScript.
	Public Property Get HttpOnly()
		HttpOnly = pHttpOnly
	End Property

	' Integer or None. Netscape cookies have version 0. RFC 2965 and RFC 2109 cookies have
	' a version cookie-attribute of 1. However, note that cookielib may ‘downgrade’ RFC 2109
	' cookies to Netscape cookies, in which case version is 0.
	Public Property Get Version()
		Version = pVersion
	End Property

	' Constructor for the cookie
	Public Default Function Make(name, value, domain, path, expires, comment, secure, http_only, version)
		pName = name
		pValue = value
		pDomain = domain
		pPath = path
		pExpires = expires
		pComment = comment
		pSecure = secure
		pHttpOnly = http_only
		pVersion = version
		Set Make = Me		
	End Function

	' This method is used to test a cookie against a given HTTP request scenario, in order to
	' tell whether the cookie should be sent in this request or not. The method has the following
	' syntax and parameters: Match(mixed $uri, [boolean $matchSessionCookies,
	' [int $now]]);
	' cookie domain: sub.example
	' request domain: sub.sub.example
	' In this example, the request would be sent because
	Public Function Match(strURL)
		Dim cookieSub, cookieDomain		

		Dim URL, domainMatches, m
		Set URL = New clsURL

		URL.FromString(strURL)

		With pDomainRegex
			.Pattern = "(?:([a-z0-9\.\-]*)\.)?((?!com)[a-z0-9\-]{3,}(?=\.[a-z\.]{2,}))\.(?:([a-z\.]{2,})$)"
			Set domainMatches = .Execute(pDomain)
		End With

		cookieSub = domainMatches.Item(0).Submatches.Item(0)
		cookieDomain = domainMatches.Item(0).Submatches.Item(1) & "." & domainMatches.Item(0).Submatches.Item(2)

		If cookieDomain = URL.Domain & "." & URL.TLD Then
			If cookieSub <> "" Then
				If Left(cookieSub, 1) = "." Then
					cookieSub = Right(cookieSub, Len(cookieSub) - 1)
					' WScript.Echo cookieSub
					' WScript.Echo URL.Subdomain
				End If

				If cookieSub = URL.Subdomain Then
					WScript.Echo "Cookie is valid!"
				ElseIf cookieSub = Right(URL.Subdomain, Len(cookieSub)) And Right(Left(URL.Subdomain, Len(URL.Subdomain) - Len(cookieSub)), 1) = "." Then
					WScript.Echo "Cookie is valid!"
				Else
					WScript.Echo "Cookie is invalid..."
				End If
			End If
		End If 

		Match = True
	End Function

	' If you just want the attributes other than name and value, you can export those to a
	' dict with the attributes() method, which produces a mapping of attribute names to encoded
	' values and is also used internally for rendering:
	' >>> cookie.attributes()
	' {'Comment': 'no'}
	Public Function Attributes()

	End Function

	' Check whether the cookie is expired or not. If the cookie has no expiration time,
	' it will always return TRUE.
	Public Function IsExpired()

	End Function

	' A cookie object can be transferred back into a string.
	' Return a string representing the Cookie, without any surrounding HTTP or JavaScript.
	Public Function ToString()

	End Function

	' Method to parse a cookie header string into a Cookie object.
	' Cookie string should be represented in the 'Set-Cookie ' HTTP response header
	' or 'Cookie' HTTP request header.
	' $cookieStr: a cookie string as represented in the 'Set-Cookie' HTTP response header or
	' 'Cookie' HTTP request header (required)
	' $refUri: a reference URI according to which the cookie's domain and path will be set.
	' (optional, defaults to parsing the value from the $cookieStr)	
	' $encodeValue: If the value should be passed through urldecode. Also effects the cookie's
	' behavior when being converted back to a cookie string. (optional, defaults to true)
	Public Function FromString(cookie_str)

	End Function

	' You can also do the reverse operation with to_dict():
	' e.g.:
	' >>> cookie = Cookie('x', 'y', comment='no')
	' >>> sorted(cookie.to_dict().items())
	' [('Comment', 'no'), ('name', 'x'), ('value', 'y')]
	Public Function ToDict()

	End Function

	' a dict that maps attribute names to values. This will parse the values as strings,
	' which can be convenient when you don’t have an existing string to parse.
	' e.g. cookie = Cookie.from_dict({'name': 'x', 'value': 'y', 'expires': 'Thu, 23 Jan 2003
	' 00:00:00 GMT'})
	Public Function FromDict(cookie_dict)

	End Function

	' Function to parse a cookie from a paramarray
	Public Function FromArray()

	End Function

	' Possibly include a FromRequest() method to parse a request into a cookie object

	'  'Set-Cookie ' HTTP response header
	' This method will produce a HTTP response "Set-Cookie" header string, showing the cookie's name
	' and value, as well as any attributes assigned to the Cookie object.
	' suitable to be sent as an HTTP header. By default, all the attributes are included, unless attrs
	' is given, in which case it should be a list of attributes to use. header is by default
	' "Set-Cookie:".
	' SetCookieHeader()

	' *** Likely do not need this method because all cookies are sent in one header for an
	' *** HTTP request from the client.
	' 'Cookie' HTTP request header.
	' This method will produce a HTTP request "Cookie" header string, showing the cookie's name
	' and value, and terminated by a semicolon (';'). The value will be URL encoded, as expected
	' in a Cookie header:
	' CookieHeader()

	Private Sub Class_Terminate()
		Set pDomainRegex = Nothing
	End Sub
End Class

If WScript.ScriptName = "v_HTTP_Cookie.vbs" Then
	Dim cookieSub, requestSub

	cookieSub = "sub.startech.com"
	requestSub = "www.sub.sub.startech.com"

	' WScript.Echo Left(requestSub, Len(requestSub) - Len(cookieSub))

	' WScript.Echo Right(Left(requestSub, Len(requestSub) - Len(cookieSub)), 1)

	' If cookieSub = Right(requestSub, Len(cookieSub)) and Right(Left(requestSub, Len(requestSub) - Len(cookieSub)), 1) = "."
	'	"Cookie is valid!"
	' Else
	'	"Cookie is not valid..."
	' End If

	Dim http_cookie
	Set http_cookie = (New clsCookie)("testCookie", "testValue", ".sub.startech.com", "/", "Wednesday, 09-Nov-99 23:12:40 GMT", "", False, True, 1)

	If http_cookie.Match("http://www.startech.com") Then
		' WScript.Echo "Cookie is valid and should be sent!"
	Else
		' WScript.Echo "Cookie is not valid. Do not send with this request"
	End If
End If