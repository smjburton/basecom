Option Explicit

Include "base_Sys_ErrorHandler"
Include "base_Sys_EventHandler"
Include "base_Sys_Info"

Const STD_IN 	= 0
Const STD_OUT 	= 1
Const STD_ERR 	= 2

Class base_Sys
	Private p_objErrorHandler, _
		p_objEventHandler, _
		p_objSysInfo

	Private Sub Class_Initialize()
		Set p_objErrorHandler = New base_Sys_ErrorHandler
		Set p_objEventHandler = New base_Sys_EventHandler
		Set p_objSysInfo = New base_Sys_Info
	End Sub


	' Properties

	
	Public Property Get ErrorHandler()
		Set ErrorHandler = p_objErrorHandler
	End Property

	Public Property Get EventHandler()
		Set EventHandler = p_objEventHandler
	End Property

	Public Property Get Info()
		Set Info = p_objSysInfo
	End Property
	

	' Methods


	Sub Quit( _
    		ByVal intExitCode _
    		)
    
    		WScript.Quit intExitCode
	End Sub

	Sub Run( _
		ByVal strScript _
		)
		On Error Resume Next

		If InStr(strScript, " ") > 0 Then
			ExecuteGlobal strScript
		Else
			Me.WriteLn CStr(Eval(strScript))
		End If

		If Err Then Call Me.ErrorHandler("base_Sys.Run")
	End Sub

	Sub Sleep( _
    		ByVal intTimeSeconds _
    		)

    		WScript.Sleep Int(intTimeSeconds * 1000)
	End Sub

	Sub Write( _
    		ByVal strText _
    		)

    		Dim objFso, _
        		objStdOut

    		Set objFso = CreateObject("Scripting.FileSystemObject")
    		Set objStdOut = objFso.GetStandardStream(STD_OUT)

    		With objStdOut
        		.Write strText
        		.Close
    		End With

    		Set objFso = Nothing
    		Set objStdOut = Nothing
	End Sub

	Sub WriteLn( _
		ByVal strText _
		)

		Dim objFso, _
			objStdOut

		Set objFso = CreateObject("Scripting.FileSystemObject")
		Set objStdOut = objFso.GetStandardStream(STD_OUT)

		With objStdOut
			.WriteLine strText
			.Close
		End With

		Set objFso = Nothing
		Set objStdOut = Nothing
	End Sub

	Private Sub Class_Terminate()
		Set p_objErrorHandler = Nothing
		Set p_objEventHandler = Nothing
		Set p_objSysInfo = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Sys.vbs" Then

End If