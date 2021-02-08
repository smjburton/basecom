Option Explicit

' “version-independent” ProgID for MSXML:
' Msxml2.ServerXMLHTTP = IServerXMLHTTPRequest object

' Msxml2.ServerXMLHTTP.3.0 = IServerXMLHTTPRequest object
' Msxml2.ServerXMLHTTP.4.0 = (?)
' Msxml2.ServerXMLHTTP.6.0 = IServerXMLHTTPRequest2 object

Class base_HTTP_XMLHTTP_ServerRequest
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

If WScript.ScriptName = "base_HTTP_XMLHTTP_ServerRequest.vbs" Then

End If