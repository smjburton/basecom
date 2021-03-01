Option Explicit

Class base_Sys_Event
	Private p_objEventHandler, _
		p_strEventType, _
		p_objEventArgs, _
		p_varCaller

	Private Sub Class_Initialize()

	End Sub

	Public Property Get Args()
		Set Args = p_objEventArgs
	End Property

	Public Property Get Caller()
		If IsObjecT(varCaller) Then
			Caller = p_varCaller
		Else
			Set Caller = p_varCaller
		End If
	End Property

	Public Property Get Handler()
		Set Handler = p_objEventHandler
	End Property

	Public Property Set Handler(objEventHandler)
		Set p_objEventHandler = objEventHandler
	End Property

	Public Property Get EventType()
		EventType = p_strEventType
	End Property

	Public Default Function Fire()
		If p_objEventArgs Then
			EventHandler(Me)
		Else
			EventHandler(Me)(p_objEventArgs)
		End If
	End Function

	Private Sub Class_Terminate()

	End Sub
End Class

If WScript.ScriptName = "base_Sys_Event.vbs" Then

End If
