Option Explicit

' The core Response object. All Request objects contain a response attribute, which is
' an instance of this class.

' See: http://framework.zend.com/manual/current/en/modules/zend.http.response.html

' Body
' RawBody
' Headers
' RawHeaders
' Request
' Status Code
' Content Type
' Parent Type
' Charset
' Meta Data
' Is Mime Vendor Specific
' Is Mime Personal

Class base_HTTP_Response
	Private Sub Class_Initialize()

	End Sub

	' Dictionary of configurations for this request.
	Public Property Get Config()

	End Property

	' Content of the response, in bytes.
	Public Property Get Content()

	End Property

	' A dictionary of Cookies the server sent back.
	Public Property Get Cookies()

	End Property

	' Encoding to decode with when accessing r.content.
	Public Property Get Encoding()

	End Property

	' Resulting HTTPError of request, if one occurred.
	Public Property Get Error()

	End Property

	' Case-insensitive Dictionary of Response Headers. For example, headers['content-encoding']
	' will return the value of a 'Content-Encoding' response header.
	Public Property Get Headers()

	End Property

	' A list of Response objects from the history of the Request. Any redirect responses
	' will end up here.
	Public Property Get History()

	End Property

	' File-like object representation of response (for advanced usage).
	Public Property Get Raw()

	End Property

	' The Request that created the Response.
	Public Property Get Request()

	End Property

	' Integer Code of responded HTTP Status.
	Public Property Get Status()

	End Property

	' Content of the response, in unicode.
	' If Response.encoding is None and chardet module is available, encoding will be guessed.
	Public Property Get Text()

	End Property

	' Final URL location of Response.
	Public Property Get URL()

	End Property

	' iter_content(chunk_size=10240, decode_unicode=False)
	' Iterates over the response data. This avoids reading the content at once into memory
	' for large responses. The chunk size is the number of bytes it should read into memory.
	' This is not necessarily the length of each item returned as decoding can take place.

	' iter_lines(chunk_size=10240, decode_unicode=None)
	' Iterates over the response data, one line at a time. This avoids reading the content at
	' once into memory for large responses.

	' raise_for_status(allow_redirects=True)
	' Raises stored HTTPError or URLError, if one occurred.

	Private Sub Class_Terminate()

	End Sub
End Class

If WScript.ScriptName = "base_HTTP_Response.vbs" Then
	Dim httpResp
	Set httpResp = New base_HTTP_Response
End If
