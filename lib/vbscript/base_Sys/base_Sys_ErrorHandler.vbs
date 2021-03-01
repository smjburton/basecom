Option Explicit

Include "base_Data.base_Data_Dictionary"
Include "base_Sys.base_Sys_Error"

Sub DefaultErrorHandler()
	PrintLn "Unhandled Error " & objError.Number & ": " & objError.Description & " (Source: " & objError.Source & ") (Procedure: " & strMethodCaller & ")" 
	Me.ReRaise objError
End Sub

Class base_Sys_ErrorHandler
	Private p_objHandlerDict

	Private Sub Class_Initialize()
		Set p_objHandlerDict = New base_Data_Dictionary
	End Sub

	Public Sub Register( _
		strMethodCaller, _
		strErrorHandler _
		)
			
		PrintLn "Registered error handler for: " & strMethodCaller

		p_objHandlerDict.Add strMethodCaller, GetRef(strErrorHandler)
	End Sub

	Public Default Function Handle( _
		strMethodCaller _
		)

		If Err.Number = 0 Then Exit Function

		Dim objError, _
			strErrorMsg

		Set objError = (New base_Sys_Error)(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
		
		On Error Resume Next

		strErrorMsg = "Error " & objError.Number & ": " & objError.Description & " (Source: " & objError.Source & ") (Procedure: " & strMethodCaller & ")"

		PrintLn strErrorMsg

		' If strMethodCaller in p_objHandlerDict and the associated item is a method,
		' call the error handler function	
		If p_objHandlerDict.Exists(strMethodCaller) Then
			Set Handle = p_objHandlerDict(strMethodCaller)
		Else
			Set Handle = GetRef("DefaultErrorHandler")
		End If
	End Function

	Public Sub Log()

	End Sub

	Public Sub Raise( _
		intErrorNumber, _
		strErrorSource, _
		strErrorDescription, _
		strErrorHelpFile, _
		strErrorHelpContext _
		)

		Err.Raise intErrorNumber, strErrorSource, strErrorDescription, strErrorHelpFile, strErrorHelpContext
	End Sub

	Sub ReRaise( _
		ByVal objError _
		)

		Err.Raise objError.Number, objError.Source, objError.Description, objError.HelpFile, objError.HelpContext
	End Sub

	Public Sub Deregister( _
		strErrorHandlerName _
		)

	End Sub

	Private Sub Class_Terminate()
		Set p_objHandlerDict = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Sys_Error.vbs" Then

End If