Option Explicit

Class base_IO_File
	Private p_File

	Private Sub Class_Initialize()
		Set p_File = CreateObject("Scripting.FileSystemObject")
	End Sub


	' Properties


	' Methods


	Public Function CreateTextFile(strFileName) ' Optional params: [Overwrite As Boolean = True], [Unicode As Boolean = False]) As TextStream
		Set CreateTextFile = p_File.CreateTextFile(strFileName) 
	End Function

	Public Function CreateTempFile(strFolder)
		Dim strTmpFileName
		strTmpFileName = Me.strTmpFileName()
		Set CreateTempFile = p_File.CreateTextFile(strFolder & "\\" & strTmpFileName)
	End Function

	Public Function FileExists(strFileSpec)
		FileExists = p_File.FileExists(strFileSpec)
	End Function

	Public Function GetFile(strFilePath)
		Set GetFile = p_File.GetFile(strFilePath)
	End Function

	Public Function GetFileVersion(strFileName)
		GetFileVersion = p_File.GetFileVersion(strFileName)
	End Function
  
	Public Function GetTempName()
		GetTempName = p_File.GetTempName()
	End Function
  
	Function OpenTextFile(strFileName) ' Optional params: [IOMode As IOMode = ForReading], [Create As Boolean = False], [Format As Tristate = TristateFalse]) As TextStream
		Set OpenTextFile = p_File.OpenTextFile(strFileName)
	End Function

	Private Sub Class_Terminate()
		Set p_File = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_File.vbs" Then

End If
