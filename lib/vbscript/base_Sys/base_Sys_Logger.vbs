Option Explicit

Include "base_IO_TextStream"

Class base_Sys_Logger
	Private p_objLogger

	Private Sub Class_Initialize()
		Set p_objLogger = New base_IO_TextStream
	End Sub

	Private Sub Class_Terminate()
		Set p_objLogger = Nothing
	End Sub
End Class