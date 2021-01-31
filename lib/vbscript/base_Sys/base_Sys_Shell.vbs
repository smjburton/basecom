Option Explicit

Class base_Shell
	Private p_FSO

	Private Sub Class_Initialize()
		Set p_FSO = CreateObject("Scripting.FileSystemObject")
	End Sub


	' Methods
	' FileSystemObject Methods

	' Public Sub CopyFile(Source As String, Destination As String, [OverWriteFiles As Boolean = True])
	' Public Sub CopyFolder(Source As String, Destination As String, [OverWriteFiles As Boolean = True])
	' Public Sub DeleteFile(FileSpec As String, [Force As Boolean = False])
	' Public Sub DeleteFolder(FolderSpec As String, [Force As Boolean = False])
	' Public Sub MoveFile(Source As String, Destination As String)
	' Public Sub MoveFolder(Source As String, Destination As String)  

	Private Sub Class_Terminate()
		Set p_FSO = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Shell.vbs" Then

End If
