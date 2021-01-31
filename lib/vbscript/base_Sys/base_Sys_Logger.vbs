Option Explicit

Include "base_IO.base_File"

Class base_Sys_Logger
	Private p_Logger

	Private Sub Class_Initialize()
		Set p_Logger = New base_File
	End Sub

	Private Sub Class_Terminate()
		Set p_Logger = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_File.vbs" Then

End If