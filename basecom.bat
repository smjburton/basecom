<!-- : Begin batch script
@echo off
cscript //nologo "%~f0?.wsf" %*
exit /b

----- Begin wsf script --->
<job>
	<script language="VBScript">
	Option Explicit

	Sub Include( _
    		ByVal strFile _
    	)

    		On Error Resume Next

    		Dim objFSO, _
        		strBasecomDirectory, _
			strFilePath

    		Set objFSO = CreateObject("Scripting.FileSystemObject")

		If InStr(strFile, ".") > 0 Then
			Dim arrLibrary
			arrLibrary = Split(strFile, ".")
			
			strBasecomDirectory = Mid(WScript.ScriptFullName, 1, InStrRev(WScript.ScriptFullName, "\"))
    			strFilePath = strBasecomDirectory & "lib\vbscript\" & arrLibrary(0)			
			strFile = arrLibrary(1)
		Else
			strFilePath = objFSO.GetAbsolutePathName(".")
 		End If   

    		ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(strFilePath & "\" & strFile & ".vbs", 1).ReadAll()

    		If Err.Number <> 0 Then
        			If Err.Number = 1041 Then 
            			Err.Clear
        		Else
            			WScript.Echo Err.Number & ": " & Err.Description
            			WScript.Quit 1
        		End If
    		End If
	End Sub

	Sub RunInteractiveInterpreter()
		Dim objStdInput, _
             	strPrompt, _
             	strInput

		Set objStdInput = WScript.StdIn
		strPrompt = ">> "

		PrintLn "basecom 0.1 (v0.1; January 23, 2021)"
		PrintLn "Type ""help"", ""copyright"", ""credits"", ""license"" for more information."
		Print strPrompt

		Do Until objStdInput.AtEndOfStream
			strInput = objStdInput.ReadLine()

			If LCase(strInput) = "exit" Or _
                    		LCase(strInput) = "exit()" Or _
                    		LCase(strInput) = "quit" Or _
                    		LCase(strInput) = "quit()" Then
					Exit Do
			Else
				Run strInput
				Print strPrompt
			End If
		Loop
	End Sub

    	Include "base_Sys.base_Sys"

        ' Options:
        ' new
        ' run
	' schedule
        ' version
        ' help
        ' credits
        ' license
        ' copyright

	If WScript.Arguments.Count > 0 Then
		Include WScript.Arguments(0)
	Else
		RunInteractiveInterpreter
	End If
	</script>
</job>