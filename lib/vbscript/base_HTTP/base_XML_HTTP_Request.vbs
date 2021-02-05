Option Explicit

' “version-independent” ProgID for MSXML:
' Microsoft.XMLHTTP = IXMLHTTPRequest object
' Microsoft.XMLHTTP.1.0 = IXMLHTTPRequest object

' “version-independent” ProgID for MSXML:
' Msxml2.XMLHTTP = IXMLHTTPRequest object

' Msxml2.XMLHTTP.3.0 = IXMLHTTPRequest object
' Msxml2.XMLHTTP.4.0 = (?) - do not use; contains issues
' Msxml2.XMLHTTP.6.0 = IServerXMLHTTPRequest2 object

Class v_XMLHTTP_Request
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

If WScript.ScriptName = "v_XMLHTTP_Request.vbs" Then

End If