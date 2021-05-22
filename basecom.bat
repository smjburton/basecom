<!-- : Begin batch script
@echo off
C:\Windows\SysWOW64\cscript //nologo "%~f0?.wsf" %*
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

		If InStr(strFile, "_") > 0 Then
			Dim arrLibrary
			arrLibrary = Split(strFile, "_")
			
			' "base_Data_Array"

			strBasecomDirectory = Mid(WScript.ScriptFullName, 1, InStrRev(WScript.ScriptFullName, "\"))
    			strFilePath = strBasecomDirectory & "lib\vbscript\" & arrLibrary(0) & "_" & arrLibrary(1)		
			' strFile = strFile & ".vbs"

			' WScript.Echo "basecom Directory: " & strBasecomDirectory 
			' WScript.Echo "File path: " & strFilePath 
			' WScript.Echo "File: " & strFile 

			ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(strFilePath & "\" & strFile & ".vbs", 1).ReadAll()
		Else
			Err.Raise 1111, "Include(..)", "Unable to load basecom library file: " & strFile, "", ""
			' Else
			'	strFilePath = objFSO.GetAbsolutePathName(".")
 		End If   

    		' ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(strFilePath & "\" & strFile & ".vbs", 1).ReadAll()

    		If Err.Number <> 0 Then
			' Library has already been included (Error 1041: Name Redefined)
        		If Err.Number = 1041 Then
				Err.Clear
			Else
            			WScript.Echo "Error " & Err.Number & ": " & Err.Description & " (Source: " & Err.Source & ")"
            			' WScript.Quit 1
			End If
    		End If
	End Sub

    	Include "base_Sys"
    	Include "base_Sys_Util"

	Sub RunDbShell()
		Include "base_Database.base_Database_Connection"

		Dim objStdInput, _
             	strPrompt, _
             	strInput

		Dim objConnection, _
			objCursor, _
			strConnectionString

		Set objStdInput = WScript.StdIn
		strPrompt = "dbshell> "

		Print strPrompt

		' Commands:
		' .open
		' .help
		' .logging
		' .quit
		' .exit
		' .clone
		' .databases
		' .tables

		strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\sburt\Desktop\Product Database.accdb;Persist Security Info=False;"

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

	If WScript.Arguments.Count > 0 Then
		If LCase(WScript.Arguments(0)) = "dbshell" Then
			RunDbShell
		Else
			Include WScript.Arguments(0)
		End If
	Else
		RunInteractiveInterpreter
	End If
	</script>
</job>