Option Explicit

Class base_Fax_Server
	Private p_FaxServer

	Private Sub Class_Initialize()
		Set p_FaxServer = CreateObject("FaxComEx.FaxServer.1")
	End Sub


	' Properties




	' Methods



	Private Sub Class_Terminate()
		Set p_FaxServer = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Fax_Server.vbs" Then

End If