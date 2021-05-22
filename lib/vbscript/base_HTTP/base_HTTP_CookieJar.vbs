Option Explicit

' A server, when returning an HTTP object to a client, may also send a piece of state information which
' the client will store. Included in that state object is a description of the range of URLs for which
' that state is valid. Any future HTTP requests made by the client which fall in that range will
' include a transmittal of the current value of the state object from the client back to the server.
' The state object is called a cookie, for no compelling reason.

' When requesting a URL from an HTTP server, the browser will match the URL against all cookies
' and if any of them match, a line containing the name/value pairs of all matching cookies will
' be included in the HTTP request. Here is the format of that line:
' Cookie: NAME1=OPAQUE_STRING1; NAME2=OPAQUE_STRING2 ...

' This is generally the format of a cookie:
' Set-Cookie: NAME=VALUE; expires=DATE;
' path=PATH; domain=DOMAIN_NAME; secure

' See: http://curl.haxx.se/rfc/cookie_spec.html

' Zend_Http_CookieJar is an object usually used by Zend_Http_Client to hold a set of
' Zend_Http_Cookie objects. The idea is that if a Zend_Http_CookieJar object is attached to
' a Zend_Http_Client object, all cookies going from and into the client through HTTP requests
' and responses will be stored by the CookieJar object. Then, when the client will send another
' request, it will first ask the CookieJar object for all cookies matching the request. These
' will be added to the request headers automatically. This is highly useful in cases where you
' need to maintain a user session over consecutive HTTP requests, automatically sending the
' session ID cookies when required. Additionally, the Zend_Http_CookieJar object can be
' serialized and stored in $_SESSION when needed.

' Once the Set-Cookie headers are received by the client, it will return those cookies
' to the server on subsequent requests using the Cookie header. The incoming header will look like:
' Cookie: integer=5; string_with_quotes="He said, \"Hello, World!\""

' CookieJar objects support the iterator protocol for iterating over contained Cookie objects.

Include "base_Data_Array_Util"
Include "base_HTTP_Cookie"

Class base_HTTP_CookieJar
	Private pCookies, _
		pCount, _
		pFile

	Private Sub Class_Initialize()
		Set pCookies = CreateObject("Scripting.Dictionary")
	End Sub

	' *******************************************************************************************
	' Also need methods for rendering HTTP response and request headers:
	' Other than parsing strings into Cookie objects, or modifying them, you might also
	' want to generate rendered output. For this, use render_request() or render_response(),
	' depending on the sort of headers you want to render. You can render all the headers at
	' once - either as separate lines, or all on one line.
	' >>> cookies.render_request()
	' 'dad=pretty; mom=strong'
	' Each individual cookie can be rendered either in the format for an HTTP request, or the
	' format for an HTTP response. Attribute values can be manipulated in natural ways and the
	' rendered output changes appropriately; but rendered request headers don’t include attributes
	' (as they shouldn’t).
	' These methods should also should also take into account domain and path attributes to know
	' whether to send a cookie or not.
	' *******************************************************************************************



	' you can still use 3 provided methods to fetch cookies from the jar object: getCookie(),
	' getAllCookies(), and getMatchingCookies(). Additionnaly, iterating over the CookieJar will
	' let you retrieve all the Zend_Http_Cookie objects from it.



	' Zend_Http_CookieJar->getCookie($uri, $cookie_name[, $ret_as]): Get a single cookie from the
	' jar, according to its URI (domain and path) and name. $uri is either a string or a
	' Zend_Uri_Http object representing the URI. $cookie_name is a string identifying the cookie
	' name. $ret_as specifies the return type as described above. $ret_type is optional, and
	' defaults to COOKIE_OBJECT.
	Public Default Property Get Cookie(name)
		WScript.Echo TypeName(pCookies(name))
		Set Cookie = pCookies(name)
	End Property

	' Zend_Http_CookieJar->getAllCookies($ret_as): Get all cookies from the jar. $ret_as specifies
	' the return type as described above. If not specified, $ret_type defaults to COOKIE_OBJECT.
	' *** Refer to JSONObj.cls file for an example.
	Public Property Get Cookies()

	End Property

	Public Property Get Count()
		Count = pCookies.Count
	End Property

	' Filename of default file in which to keep cookies. This attribute may be assigned to.
	Public Property Get File()

	End Property

	' Add a cookie to the jar.
	' Zend_Http_CookieJar->addCookie($cookie[, $ref_uri]): Add a single cookie to the jar. $cookie
	' can be either a Zend_Http_Cookie object or a string, which will be converted automatically
	' into a Cookie object. If a string is provided, you should also provide $ref_uri - which is
	' a reference URI either as a string or Zend_Uri_Http object, to use as the cookie's default
	' domain and path.
	' Should be able to add a cookie whether it's a string, dictionary, or Cookie object
	Public Sub Add(http_cookie)
		Dim c

		If TypeName(http_cookie) = "clsCookie" Then
			Set c = http_cookie
		Else
			Set c = New clsCookie
			If TypeName(http_cookie) = "String" Then
				c.FromString(http_cookie)
			ElseIf TypeName(http_cookie) = "Dictionary" Then
				c.FromDict(http_cookie)
			Else
				' Error
			End If
		End If

		pCookies.Add c.Name, c
	End Sub

	' Removes a cookie from the jar
	Public Sub Remove()

	End Sub

	' Extract cookies from HTTP response and store them in the CookieJar, where allowed by policy.
	' The extract_cookies() method will look for Set-Cookie: and Set-Cookie2: headers in the
	' HTTP::Response object passed as argument. Any of these headers that are found are used to
	' update the state of the $cookie_jar.
	Public Function ExtractCookies()

	End Function

	' Load cookies from a file.
	' Old cookies are kept unless overwritten by newly loaded ones.
	Public Sub Load(filename)

	End Sub

	' Save cookies to a file.
	' filename is the name of file in which to save cookies. If filename is not specified, 
	' self.filename is used (whose default is the value passed to the constructor, if any);
	' The file is overwritten if it already exists, thus wiping all the cookies it contains.
	' Saved cookies can be restored later using the load() or revert() methods.
	Public Sub Save(filename)

	End Sub

	' This method empties the CookieJar and re-loads the CookieJar from the last save file.
	Public Sub Revert()

	End Sub

	' Empties the cookie jar.
	' Potentially add in options for selecting and clearing cookies based on domain, path, and key
	' Invoking this method without arguments will empty the whole $cookie_jar. If given a single
	' argument only cookies belonging to that domain will be removed. If given two arguments, cookies
	' belonging to the specified path within that domain are removed. If given three arguments, then
	' the cookie with the specified key, path and domain is removed.
	Public Sub Clear()

	End Sub

	' Discard all temporary cookies. Scans for all cookies in the jar with either no expire
	' field or a true discard flag. To be called when the user agent shuts down according to RFC 2965.
	Public Sub ClearTemporary()

	End Sub

	' Should consider a boolean 'Contains()' method for checking the presence of cookies

	' Need a match method to return all cookies that match a particular domain and path
	' Zend_Http_CookieJar->getMatchingCookies($uri[, $matchSessionCookies[, $ret_as[, $now]]]):
	' Get all cookies from the jar that match a specified scenario, that is a URI and expiration time.
	' This method selects and returns cookies within the CookieJar that match the domain, path, and
	' key specified
	Public Sub Match()

	End Sub

	' This method will return the state of the $cookie_jar represented as a sequence of "Set-Cookie3"
	' header lines separated by "\n". If $skip_discardables is TRUE, it will not return lines for cookies
	' with the Discard attribute.
	Public Function ToString()

	End Function

	' Another way to instantiate a CookieJar object is to use the static
	' Zend_Http_CookieJar::fromResponse() method. This method takes two parameters: a
	' Zend_Http_Response object, and a reference URI, as either a string or a Zend_Uri_Http object.
	' This method will return a new Zend_Http_CookieJar object, already containing the cookies set
	' by the passed HTTP response. The reference URI will be used to set the cookie's domain and path,
	' if they are not defined in the Set-Cookie headers.
	' Zend_Http_CookieJar->addCookiesFromResponse($response, $ref_uri): Add all cookies set in a
	' single HTTP response to the jar. $response is expected to be a Zend_Http_Response object with
	' Set-Cookie headers. $ref_uri is the request URI, either as a string or a Zend_Uri_Http object,
	' according to which the cookies' default domain and path will be set.
	Public Function FromResponse()

	End Function

	Private Sub Class_Terminate()
		Set pCookies = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_HTTP_CookieJar.vbs" Then
	Dim cj, c1, c2, c3
	Set cj = New clsCookieJar
	Set c1 = New clsCookie
	' Set c3 = CreateObject("Scripting.Dictionary")

	c1.Make "testCookie", _
		"testValue", _
		"www.startech.com", _
		"/", _
		"Wednesday, 09-Nov-99 23:12:40 GMT", _
		"", _
		False, _
		True, _
		1

	' c2 = "Set-Cookie: PART_NUMBER=ROCKET_LAUNCHER_0001; path=/"
	' c3.Add "enwiki_session", "17ab96bd8ffbe8ca58a78657a918558"

	cj.Add c1
	' cj.Add c2
	' cj.Add c3

	WScript.Echo cj("testCookie").Value
End If
