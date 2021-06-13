Option Explicit

Include "base_IO_TextStream"

Private Const FOR_WRITING = 2
Private Const FOR_APPENDING = 8

Class base_CSV_Writer
	Private p_CsvWriter, _
		p_Delimiter, _
		p_EscapeChar, _
		p_Filename, _
		p_LineTerminator, _
		p_QuoteChar, _
		p_Quoting, _
		p_SkipInitialSpace

	Private Sub Class_Initialize()
		Set p_CsvWriter = New base_IO_TextStream
		p_Delimiter = ","
		p_EscapeChar = ""
		p_Filename = ""
		p_LineTerminator = vbCrLf
		p_QuoteChar = """"
		p_Quoting = ""
		p_SkipInitialSpace = False
	End Sub


	' Properties


	Public Property Get Delimiter()
		Delimiter = p_Delimiter
	End Property

	Public Property Let Delimiter( _
		ByVal strDelimiter _
		)
		p_Delimiter = strDelimiter
	End Property

	Public Property Get EscapeChar()
		EscapeChar = p_EscapeChar 
	End Property

	Public Property Get Filename()
		Filename = p_Filename
	End Property

	Public Property Get LineTerminator()
		LineTerminator = p_LineTerminator 
	End Property

	Public Property Get QuoteChar()
		QuoteChar = p_QuoteChar 
	End Property

	Public Property Get Quoting()
		Quoting = p_Quoting 
	End Property

	Public Property Get SkipInitialSpace()
		SkipInitialSpace = p_SkipInitialSpace 
	End Property


	' Methods


	Public Sub Close()
		p_CsvWriter.Close
	End Sub

	Public Sub Create( _
		ByVal strFileName _
		)

		p_CsvWriter.Create strFileName
		p_Filename = strFileName
	End Sub

	Public Sub Open( _
		ByVal strFileName, _
		ByVal blnOverwrite _
		)

		If blnOverwrite Then
			p_CsvWriter.Open strFileName, FOR_WRITING
		Else
			p_CsvWriter.Open strFileName, FOR_APPENDING
		End If

		p_Filename = strFileName
	End Sub

	Public Sub WriteHeader( _
		ByVal varHeader _
		) 

	End Sub

	Public Sub WriteRow( _
		ByVal varRow _
		)

		Dim strRow

		If TypeName(varRow) = "String" Then
			strRow = Replace(varRow, p_Delimiter & " ", p_Delimiter)
		ElseIf IsArray(varRow) Then
			Dim varItem
		
			For Each varItem In varRow
				strRow = strRow & varItem & p_Delimiter
			Next

			strRow = Left(strRow, Len(strRow) - 1)
		End If

		p_CsvWriter.WriteLine strRow
	End Sub

	Public Sub WriteRows( _
		ByVal arrRows _
		)

		If IsArray(arrRows) Then
			Dim varItem

			For Each varItem in arrRows
				Me.WriteRow varItem
			Next
		End If
	End Sub

	Private Sub Class_Terminate()
		Set p_CsvWriter = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_CSV_Writer.vbs" Then

End If