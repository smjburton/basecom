Option Explicit

Include "base_Sys_Folder"

' File attributes can be set using any combination of the bit values below:
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

Class base_Sys_File
	Private p_File

	Private Sub Class_Initialize()

	End Sub


	' Properties


	Public Property Get Attributes()
		Attributes = p_File.Attributes
	End Property

	Public Property Let Attributes(intAttributes)
		p_File.Attributes = intAttributes
	End Property

	Public Property Get DateCreated()
		DateCreated = p_File.DateCreated
	End Property

	Public Property Get DateLastAccessed()
		Attributes = p_File.Attributes
	End Property

	Public Property Get DateLastModified()
		DateLastModified = p_File.DateLastModified
	End Property

	Public Property Get Drive()
		Drive = p_File.Drive
	End Property

	Public Property Get Name()
		Name = p_File.Name
	End Property

	Public Property Let Name(strName)
		p_File.Name = strName
	End Property

	Public Property Get ParentFolder()
		Dim objFolder
		Set objFolder = New base_Sys_Folder
		objFolder.FromFolder p_File.ParentFolder
		Set ParentFolder = objFolder 
	End Property

	Public Property Get Path()
		Path = p_File.Path
	End Property

	Public Property Get ShortName()
		ShortName = p_File.ShortName
	End Property

	Public Property Get ShortPath()
		ShortPath = p_File.ShortPath
	End Property

	Public Property Get Size()
		Size = p_File.Size
	End Property

	Public Property Get FileType()
		FileType = p_File.Type
	End Property


	' Methods


	Public Sub Copy(strDestination) ' Optional Params: [OverWriteFiles As Boolean = True])
		p_File.Copy strDestination
	End Sub

	Public Sub Delete() ' Optional Params: [Force As Boolean = False])
		p_File.Delete
	End Sub

	Public Sub Move(strDestination)
		p_File.Move strDestination
	End Sub

	Public Sub FromFile(objFile)
		Set p_File = objFile
	End Sub

	Public Sub FromPath(strFilePath)
		Dim objFileSystemObject
		Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
		Set p_File = objFileSystemObject.GetFile(strFilePath)
		Set objFileSystemObject = Nothing
	End Sub

	Private Sub Class_Terminate()
		If Not p_File Is Nothing Then Set p_File = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Sys_File.vbs" Then

End If
