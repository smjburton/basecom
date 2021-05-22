Option Explicit

Include "base_Sys_File"
Include "base_Data_Collection"

Class base_Sys_Files
	Private p_Files

	Private Sub Class_Initialize()
		Set p_Files = New base_Data_Collection
	End Sub


	' Properties


	Public Property Get Count()
		Count = p_Files.Count
	End Property

	Public Default Property Get File(intIndex)
		Set File = p_Files(intIndex)
	End Property


	' Methods


	Public Sub FromFiles(objFiles)
		Dim objFile, _
			objSysFile

    		For Each objFile in objFiles
			Set objSysFile = New base_Sys_File
			objSysFile.FromFile objFile
			p_Files.Add objSysFile
    		Next
	End Sub

	Private Sub Class_Terminate()
		Set p_Files = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Sys_Files.vbs" Then

End If
