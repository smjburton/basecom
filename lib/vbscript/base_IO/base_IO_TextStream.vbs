Option Explicit

Class base_IO_TextStream
	Private p_FileSystemObject, _
		p_TextStream

	Private Sub Class_Initialize()
		Set p_FileSystemObject = CreateObject("Scripting.FileSystemObject")
	End Sub


	' Properties


	Public Property Get AtEndOfLine()
		AtEndOfLine = p_TextStream.AtEndOfLine
	End Property

	Public Property Get AtEndOfStream()
		AtEndOfStream = p_TextStream.AtEndOfStream
	End Property

	Public Property Get Column()
		Column = p_TextStream.Column
	End Property

	Public Property Get Line()
		Line = p_TextStream.Line
	End Property


	' Methods


	Public Sub Close()
		p_TextStream.Close
	End Sub

	Public Function CreateTextFile(strFileName) ' Optional params: [Overwrite As Boolean = True], [Unicode As Boolean = False]) As TextStream
		Set p_TextStream = p_FileSystemObject.CreateTextFile(strFileName) 
	End Function
  
	Function OpenTextFile(strFileName) ' Optional params: [IOMode As IOMode = ForReading], [Create As Boolean = False], [Format As Tristate = TristateFalse]) As TextStream
		Set p_TextStream = p_FileSystemObject.OpenTextFile(strFileName)
	End Function

	Public Function Read()
		If Not p_TextStream Is Nothing Then
			Read = p_TextStream.Read()
		Else
			Read = ""
		End If
	End Function

	Public Function ReadAll()
		If Not p_TextStream Is Nothing Then
			ReadAll = p_TextStream.ReadAll()
		Else
			ReadAll = ""
		End If
	End Function

	Public Function ReadLine()
		If Not p_TextStream Is Nothing Then
			ReadLine = p_TextStream.ReadLine()
		Else
			ReadLine = ""
		End If
	End Function

	Public Sub Skip()
		If Not p_TextStream Is Nothing Then p_TextStream.Skip
	End Sub

	Public Sub SkipLine()
		If Not p_TextStream Is Nothing Then p_TextStream.SkipLine
	End Sub

	Public Sub Write(strText)
		If Not p_TextStream Is Nothing Then p_TextStream.Write strText
	End Sub

	Public Sub WriteBlankLines(intNumBlankLines)
		If Not p_TextStream Is Nothing Then p_TextStream.WriteBlankLines intNumBlankLines
	End Sub

	Public Sub WriteLine(strText)
		If Not p_TextStream Is Nothing Then p_TextStream.WriteLine strText
	End Sub

	Private Sub Class_Terminate()
		Set p_FileSystemObject = Nothing
		If Not p_TextStream Is Nothing Then Set p_TextStream = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_IO_TextStream.vbs" Then
	Dim objTextStream
	Set objTextStream = New base_IO_TextStream

	With objTextStream
		.CreateTextFile "C:\Dev\Projects\basecom\lib\vbscript\base_IO\test.txt"
		.WriteLine "Test, test, test"
		.Close
	End With
End If
