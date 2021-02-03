Option Explicit

Sub ErrorHandler( _
	varHandlerMethod, _
	varHandlerMethodParams() _
	)

	WScript.Echo "Entered ErrorHandler()"

	If Err.Number = 0 Then Exit Sub

	Dim strErrorMessage

	strErrorMessage = "Error " & Err.Number & ": " & Err.Description & " (Source: " & Err.Source & ")"

	' Need filepath
	' Need stacktrace

	WScript.Echo strErrorMessage
        ' WScript.StdErr.WriteLine strErrorMessage

	If IsMethod(varHandlerMethod) Then Call GetRef(varHandlerMethod)(varHandlerMethodParams)
	
	' Use a global error logger condition to determine whether to write to the error log
	' If varErrorLogger Then varErrorLogger.WriteLine strErrorMessage
 
	Err.Clear
End Sub

Sub RaiseError( _
	intErrorNumber, _
	strErrorSource, _
	strErrorDescription, _
	strErrorHelpFile, _
	strErrorHelpContext _
	)

	Err.Raise intErrorNumber, strErrorSource, strErrorDescription, strErrorHelpFile, strErrorHelpContext
End Sub

Sub ReRaiseError()
	Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Sub TestErrorHandlerMethod( _
	varHandlerMethodParams() _
	)

	' Test for expected error number
	' If Err.Number <> intExpectedErrNum Then 
	' re-raise error with additional context if not handled in the error handler method
	' If blnQuitOnError Then WScript.Quit

	WScript.Echo "Error received from: " & varHandlerMethodParams(0)
End Sub

Sub TestError()
	On Error Resume Next

	Dim intNum

	intNum = 1/0
	
	If Err Then Call ErrorHandler("TestErrorHandlerMethod"), Array("TestError() Method"))
End Sub

If WScript.ScriptName = "base_Sys_Error.vbs" Then
	TestError()
End If