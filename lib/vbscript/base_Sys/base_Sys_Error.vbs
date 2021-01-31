Option Explicit

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