Option Explicit

Include "base_Sys.base_Sys_Files"
Include "base_Sys.base_Sys_Folders"

' Folder attributes can be set using a combination of the bit values below:
' Constant	Value	Description
' Normal	0	Normal file. No attributes are set.
' ReadOnly	1	Read-only file. Attribute is read/write.
' Hidden	2	Hidden file. Attribute is read/write.
' System	4	System file. Attribute is read/write.
' Volume	8	Disk drive volume label. Attribute is read-only.
' Directory	16	Folder or directory. Attribute is read-only.
' Archive	32	File has changed since last backup. Attribute is read/write.
' Alias		1024	Link or shortcut. Attribute is read-only.
' Compressed	2048	Compressed file. Attribute is read-only.

Class base_Sys_Folder
	Private p_Folder

	Private Sub Class_Initialize()

	End Sub


	' Properties:


	Public Property Get Attributes()
		Attributes = p_Folder.Attributes
	End Property

	Public Property Let Attributes(intAttributes)
		p_Folder.Attributes = intAttributes
	End Property

	Public Property Get DateCreated()
		DateCreated = p_Folder.DateCreated
	End Property

	Public Property Get DateLastAccessed()
		DateLastAccessed = p_Folder.DateLastAccessed
	End Property

	Public Property Get DateLastModified()
		DateLastModified = p_Folder.DateLastModified
	End Property

	Public Property Get Drive()
		Drive = p_Folder.Drive
	End Property

	Public Property Get Files()
		Dim objFiles
		Set objFiles = New base_Sys_Files
		objFiles.FromFiles p_Folder.Files
		Set Files = objFiles
	End Property

	Public Property Get IsRootFolder()
		IsRootFolder = p_Folder.IsRootFolder
	End Property

	Public Property Get Name()
		Name = p_Folder.Name
	End Property

	Public Property Let Name(strName)
		p_Folder.Name = strName
	End Property

	Public Property Get ParentFolder()
		ParentFolder = p_Folder.ParentFolder
	End Property

	Public Property Get Path()
		Path = p_Folder.Path
	End Property

	Public Property Get ShortName()
		ShortName = p_Folder.ShortName
	End Property

	Public Property Get ShortPath()
		ShortPath = p_Folder.ShortPath
	End Property

	Public Property Get SubFolders()
		Dim objFolders
		Set objFolders = New base_Sys_Folders
		objFolders.FromFolders p_Folder.SubFolders
		Set SubFolders = objFolders
	End Property

	Public Property Get FolderType()
		FolderType = p_Folder.Type
	End Property


	' Methods:

	
	Public Sub AddFolders(strFolderName)
		p_Folder.AddFolders strFolderName
	End Sub

	Public Sub Copy(strDestination) ' Optional Params: [OverWriteFiles As Boolean = True])
		p_Folder.Copy strDestination
	End Sub

	Public Sub Delete() ' Optional Params: [Force As Boolean = False])
		p_Folder.Delete
	End Sub

	Public Sub Move(strDestination)
		p_Folder.Move strDestination
	End Sub

	Public Sub FromFolder(objFolder)
		Set p_Folder = objFolder
	End Sub

	Public Sub FromPath(strFolderPath)
		Dim objFileSystemObject
		Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
		Set p_Folder = objFileSystemObject.GetFolder(strFolderPath)
		Set objFileSystemObject = Nothing
	End Sub

	Private Sub Class_Terminate()
		If Not p_Folder Is Nothing Then Set p_Folder = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Sys_Folder.vbs" Then

End If