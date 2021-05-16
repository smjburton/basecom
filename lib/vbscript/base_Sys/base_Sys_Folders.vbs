Option Explicit

Include "base_Sys.base_Sys_Folder"
Include "base_Data.base_Data_Collection"

Class base_Sys_Folders
	Private p_Folders

	Private Sub Class_Initialize()
		Set p_Folders = New base_Data_Collection
	End Sub


	' Properties


	Public Property Get Count()
		Count = p_Folders.Count
	End Property

	Public Default Property Get Folder(intIndex)
		Set Folder = p_Folders(intIndex)
	End Property


	' Methods


	Public Sub FromFolders(objFolders)
		Dim objFolder, _
			objSysFolder

    		For Each objFolder in objFolders
			Set objSysFolder = New base_Sys_Folder
			objSysFolder.FromFolder objFolder
			p_Folders.Add objSysFolder 
    		Next
	End Sub

	Private Sub Class_Terminate()
		Set p_Folders = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Sys_Folders.vbs" Then

End If