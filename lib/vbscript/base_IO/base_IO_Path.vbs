Option Explicit

Class base_IO_Path
	Private p_Path

	Private Sub Class_Initialize()
		Set p_Path = CreateObject("Scripting.FileSystemObject")
	End Sub


	' Properties


	Public Property Get Drives()
		Set Drives = p_File.Drives
	End Property


	' Methods


	' Public Function BuildPath(Path As String, Name As String) As String
	' Public Function DriveExists(DriveSpec As String) As Boolean
	' Public Function FileExists(FileSpec As String) As Boolean
	' Public Function FolderExists(FolderSpec As String) As Boolean    
	' Public Function GetAbsolutePathName(Path As String) As String
	' Public Function GetBaseName(Path As String) As String 
	' Public Function GetDrive(DriveSpec As String) As Drive 
	' Public Function GetDriveName(Path As String) As String 
	' Public Function GetExtensionName(Path As String) As String 
	' Public Function GetFile(FilePath As String) As File
	' Public Function GetFileName(Path As String) As String
	' Public Function GetFileVersion(Path as String) As String
	' Public Function GetFolder(FolderPath As String) As Folder 
	' Public Function GetParentFolderName(Path As String) As String
	' Public Function GetSpecialFolder(SpecialFolder As SpecialFolderConst) As Folder 


	Private Sub Class_Terminate()
		Set p_Path = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_IO_Path.vbs" Then

End If
