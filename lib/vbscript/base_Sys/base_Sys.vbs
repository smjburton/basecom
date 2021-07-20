Option Explicit

Include "base_Sys_ErrorHandler"
Include "base_Sys_EventHandler"

Dim ErrorHandler, _
	EventHandler

Set ErrorHandler = New base_Sys_ErrorHandler
Set EventHandler = New base_Sys_EventHandler

Const STD_IN 	= 0
Const STD_OUT 	= 1
Const STD_ERR 	= 2

Sub Print( _
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
End Sub

Sub PrintLn( _
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
End Sub

Sub Run( _
	ByVal strScript _
	)
	On Error Resume Next

	If InStr(strScript, " ") > 0 Then
		ExecuteGlobal strScript
	Else
		PrintLn CStr(Eval(strScript))
	End If

	If Err Then Call ErrorHandler("base_Sys.Run")
End Sub

Sub Sleep( _
    ByVal intTimeSeconds _
    )

    WScript.Sleep Int(intTimeSeconds * 1000)
End Sub

' Function FunctionExists( func_name )'
'     FunctionExists = False 
' 
'     On Error Resume Next
' 
'     Dim f : Set f = GetRef(func_name)
' 
'     If Err.number = 0 Then
'         FunctionExists = True
'     End If  
'     On Error GoTo 0
' 
' End Function

Sub Quit( _
    ByVal intExitCode _
    )
    
    WScript.Quit intExitCode
End Sub