Option Explicit

Include "base_Data.base_Data_Dictionary"
Include "base_Sys.base_Sys_Event"

Class base_Sys_EventHandler
	Private p_objEventHandlerDict

	Private Sub Class_Initialize()
		Set p_objEventHandlerDict = New base_Data_Dictionary
	End Sub

	Public Sub Register( _
		objEvent, _
		objEventHandler _
		)

		PrintLn "Registering event..."

		p_objEventHandlerDict.Add objEvent, objEventHandler
	End Sub

	Public Default Function Handle( _
		objEvent _
		)

		PrintLn "Handling event..."

		If p_objEventHandlerDict.Exists(objEvent) Then
			Set Handle = p_objEventHandlerDict(objEvent)
		End If
	End Function

	Public Sub Log()

	End Sub

	Public Sub Deregister( _
		objEvent _
		)

		p_objEventHandlerDict.Remove objEvent

		' If Not p_objEventHandlerDict Is Nothing Then p_objEventHandlerDict.Remove objEvent
	End Sub	

	Private Sub Class_Terminate()
		Set p_objEventHandlerDict = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Sys_EventHandler.vbs" Then

End If