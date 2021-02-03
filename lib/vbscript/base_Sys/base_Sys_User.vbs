Option Explicit

Class base_Sys_User
	Private Sub Class_Initialize()

	End Sub

' WScript.Network.1

' ComputerName	Property ComputerName As String
'     read-only
' UserDomain	Property UserDomain As String
'     read-only
' UserName	Property UserName As String
'     read-only

' ADSystemInfo

' ExpandEnvironmentStrings	Function ExpandEnvironmentStrings(Src As String) As String

' WScript.Shell.Environment

' Environment
' Valid parameters for Environment are PROCESS, SYSTEM, USER and VOLATILE.

	Private Sub Class_Terminate()

	End Sub
End Class

If WScript.ScriptName = "base_Sys_User.vbs" Then

End If
