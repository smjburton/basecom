Option Explicit

Class base_Sys_Event
	Private Sub Class_Initialize()

	End Sub

	Public Default Function Fire()
		EventHandler(Me)
	End Function

	Private Sub Class_Terminate()

	End Sub
End Class

If WScript.ScriptName = "base_Sys_Event.vbs" Then

End If
