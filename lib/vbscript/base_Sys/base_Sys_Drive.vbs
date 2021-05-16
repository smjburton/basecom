Option Explicit

Include "base_Sys.base_Sys_Folder"

' Drive Types:

' CONST Unknown = 0
' CONST Removable = 1
' CONST Fixed = 2
' CONST Network = 3
' CONST CD-ROM = 4
' CONST RAM Disk = 5

Class base_Sys_Drive
	Private p_Drive

	Private Sub Class_Initialize()

	End Sub


	' Properties


	Public Property Get AvailableSpace()
		AvailableSpace = p_Drive.AvailableSpace
	End Property

	Public Property Get DriveLetter()
		DriveLetter = p_Drive.DriveLetter
	End Property

	Public Property Get DriveType()
		DriveType = p_Drive.DriveType
	End Property

	Public Property Get FileSystem()
		FileSystem = p_Drive.FileSystem
	End Property

	Public Property Get FreeSpace()
		FreeSpace = p_Drive.FreeSpace
	End Property

	Public Property Get IsReady()
		IsReady = p_Drive.IsReady
	End Property

	Public Property Get Path()
		Path = p_Drive.Path
	End Property

	Public Property Get RootFolder()
		Dim objFolder
		Set objFolder = New base_Sys_Folder
		objFolder.FromFolder p_Drive.RootFolder
		Set RootFolder = objFolder
	End Property

	Public Property Get SerialNumber()
		SerialNumber = p_Drive.SerialNumber
	End Property

	Public Property Get ShareName()
		ShareName = p_Drive.ShareName
	End Property

	Public Property Get TotalSize()
		TotalSize = p_Drive.TotalSize
	End Property

	Public Property Get VolumeName()
		VolumeName = p_Drive.VolumeName
	End Property

	Public Property Let VolumeName(strName)
		p_Drive.VolumeName = strName
	End Property


	' Methods


	Public Sub FromDrive(objDrive)
		Set p_Drive = objDrive
	End Sub

	Public Sub FromPath(strPath)
		Dim objFileSystemObject
		Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
		Set p_Drive = objFileSystemObject.GetDrive(strPath)
		Set objFileSystemObject = Nothing
	End Sub

	Private Sub Class_Terminate()
		If Not p_Drive Is Nothing Then Set p_Drive = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Sys_Drive.vbs" Then

End If