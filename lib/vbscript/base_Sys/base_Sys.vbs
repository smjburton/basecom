Option Explicit

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

Sub ErrorHandler()
	If Err.Number = 0 Then Exit Sub

        ' Handle specific error
	' If not the error we were expecting, re-raise the error
	' Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

	' Need filepath
	' Boolean for quit
        ' Boolean to log the erro

        WScript.StdErr.WriteLine "Error " & Err.Number & ": " & Err.Description & " (Source: " & Err.Source & ")"
	Err.Clear
End Sub

Sub Raise()
    ' Err.Raise
End Sub

Sub Run( _
    ByVal strScript _
    )
    On Error Resume Next

    ExecuteGlobal strScript

    If Err Then Call ErrorHandler
End Sub

Sub Sleep( _
    ByVal intTimeSeconds _
    )

    WScript.Sleep Int(intTimeSeconds * 1000)
End Sub

Sub Quit( _
    ByVal intExitCode _
    )
    
    WScript.Quit intExitCode
End Sub