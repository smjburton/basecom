Option Explicit

Include "base_Sys.base_Sys_Drives"
Include "base_Sys.base_Sys_Drive"
Include "base_Sys.base_Sys_Folder"
Include "base_Sys.base_Sys_File"

' CONST SystemFolder	1	 
' CONST TemporaryFolder	2	 
' CONST WindowsFolder	0

Class base_Sys_Path
	Private p_Path

	Private Sub Class_Initialize()
		Set p_Path = CreateObject("Scripting.FileSystemObject")
	End Sub


	' Properties


	Public Property Get Drives()
		Dim objDrives
		Set objDrives = New base_Sys_Drives
		objDrives.FromDrives p_Path.Drives
		Set Drives = objDrives
	End Property


	' Methods


	Public Function BuildPath(strPath, strName)
		BuildPath = p_Path.BuildPath(strPath, strName)
	End Function

	Public Sub CopyFile(strSource, strDestination) ' Optional Params: [OverWriteFiles As Boolean = True])
		p_Path.CopyFile strSource, strDestination
	End Sub

	Public Sub CopyFolder(strSource, strDestination) ' Optional Params: [OverWriteFiles As Boolean = True])
		p_Path.CopyFolder strSource, strDestination
	End Sub

	Public Sub CreateFolder(strFolderName)
		p_Path.CreateFolder strFolderName
	End Sub

	Public Sub DeleteFile(strFileSpec) ' Optional Params: [Force As Boolean = False])
		p_Path.DeleteFile strFileSpec
	End Sub

	Public Sub DeleteFolder(strFolderSpec) ' Optional Params: [Force As Boolean = False])
		p_Path.DeleteFolder strFolderSpec
	End Sub

	Public Function DriveExists(strDriveSpec)
		DriveExists = p_Path.DriveExists(strDriveSpec)
	End Function

	Public Function FileExists(strFileSpec)
		FileExists = p_Path.FileExists(strFileSpec)
	End Function

	Public Function FolderExists(strFolderSpec)
		FolderExists = p_Path.FolderExists(strFolderSpec)
	End Function
   
	Public Function GetAbsolutePathName(strPath)
		GetAbsolutePathName = p_Path.GetAbsolutePathName(strPath)
	End Function

	Public Function GetBaseName(strPath)
		GetBaseName = p_Path.GetBaseName(strPath)
	End Function

	Public Function GetDrive(strDriveSpec)
		Dim objDrive
		Set objDrive = New base_Sys_Drive
		objDrive.FromDrive p_Path.GetDrive(strDriveSpec)
		Set GetDrive = objDrive
	End Function
 
	Public Function GetDriveName(strPath)
		GetDriveName = p_Path.GetDriveName(strPath)
	End Function

	Public Function GetExtensionName(strPath)
		GetDriveName = p_Path.GetExtensionName(strPath)
	End Function

	Public Function GetFile(strFilePath)
		Dim objFile
		Set objFile = New base_Sys_File
		objFile.FromFile p_Path.GetFile(strFilePath)
		Set GetFile = objFile
	End Function

	Public Function GetFileName(strPath)
		GetFileName = p_Path.GetFileName(strPath)
	End Function

	Public Function GetFileVersion(strPath)
		GetFileVersion = p_Path.GetFileVersion(strPath)
	End Function

	Public Function GetFolder(strFolderPath)
		Dim objFolder
		Set objFolder = New base_Sys_Folder
		objFolder.FromFolder p_Path.GetFolder(strFolderPath)
		Set GetFolder = objFolder
	End Function

	Public Function GetParentFolderName(strPath)
		GetParentFolderName = p_Path.GetParentFolderName(strPath)
	End Function

	Public Function GetSpecialFolder(intSpecialFolder)
		Dim objFolder
		Set objFolder = New base_Sys_Folder
		objFolder.FromFolder p_Path.GetSpecialFolder(intSpecialFolder)
		Set GetSpecialFolder = objFolder
	End Function

	Public Function GetTempName()
		GetTempName = p_Path.GetTempName()
	End Function

	Public Sub MoveFile(strSource, strDestination)
		p_Path.MoveFile strSource, strDestination
	End Sub

	Public Sub MoveFolder(strSource, strDestination)    
		p_Path.MoveFolder strSource, strDestination
	End Sub

	Private Sub Class_Terminate()
		Set p_Path = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Sys_Path.vbs" Then

End If
