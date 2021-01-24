Option Explicit

' Checks to ensure we are running the 64-bit version of Windows Script Host
If InStr(1, WScript.FullName, "system32", vbTextCompare) > 0 And CreateObject("Scripting.FileSystemObject").FileExists("C:\Windows\SysWow64\WScript.exe") = True Then
	CreateObject("WScript.Shell").Run "C:\Windows\SysWow64\WScript.exe" & " """ & WScript.ScriptFullName & """", 1, False
	WScript.Quit
End If

' Defines the Include function
Sub Include( _
    ByVal strFile _
    )

    On Error Resume Next

    Dim objFSO, _
        strFilePath

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strFilePath = objFSO.GetAbsolutePathName(".")

    ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(strFilePath & "\basecom\lib\vbscript\base_Database\" & strFile & ".vbs", 1).ReadAll()

    If Err.Number <> 0 Then
        If Err.Number = 1041 Then 
            Err.Clear
        Else
            WScript.Echo Err.Number & ": " & _
                        Err.Description
            WScript.Quit 1
        End If
    End If
End Sub

If InStr(1, WScript.FullName, "WScript", vbTextCompare) > 0 Then
	CreateObject("WScript.Shell").Run "C:\Windows\SysWow64\CScript.exe" & " """ & WScript.ScriptFullName & """", 1, False
	WScript.Quit
End If

If WScript.ScriptName = "main_template.vbs" Then
	WScript.Echo "Hello"
End If