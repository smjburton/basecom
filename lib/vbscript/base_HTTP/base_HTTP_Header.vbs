Option Explicit

' See: http://framework.zend.com/manual/current/en/modules/zend.http.headers.html

Class base_HTTP_Header
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

If WScript.ScriptName = "base_HTTP_Header.vbs" Then
	Dim httpHeader
	Set httpHeader = New base_HTTP_Header
End If
