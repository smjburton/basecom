Option Explicit

Include "base_IO_TextStream"

Private Const FOR_READING = 1

Class base_CSV_Reader
	Private p_CsvReader, _
		p_Delimiter, _
		p_Filename

	Private Sub Class_Initialize()
		Set p_CsvReader = New base_IO_TextStream
		p_Delimiter = ","
		p_Filename = ""
	End Sub


	' Properties


	Public Property Get AtEndOfFile()
		AtEndOfFile = p_CsvReader.AtEndOfStream
	End Property

	Public Property Get Delimiter()
		Delimiter = p_Delimiter
	End Property

	Public Property Get Filename()
		Filename = p_Filename
	End Property

	Public Property Get Row()
		Row = Split(p_CsvReader.ReadLine(), p_Delimiter)
	End Property

	Public Property Get RowCount()
		RowCount = UBound(Split(p_CsvReader.ReadAll(), vbCrLf))
		Me.Reset()
	End Property

	Public Property Get RowNumber()
		RowNumber = p_CsvReader.Line
	End Property

	Public Default Property Get Rows()
		Dim arrRows, _
			intIndex

		Me.Reset()
		arrRows = Split(p_CsvReader.ReadAll(), vbCrLf)
		
		For intIndex = 0 To UBound(arrRows)
			arrRows(intIndex) = Split(arrRows(intIndex), p_Delimiter)
		Next

		Rows = arrRows
		Me.Reset()		
	End Property


	' Methods


	Public Sub Close()
		p_CsvReader.Close
	End Sub

	Public Sub Open( _
		ByVal strFileName _
		)

		p_CsvReader.Open strFileName, FOR_READING
		p_Filename = strFileName
	End Sub

	Public Function ReadRow()
		ReadRow = p_CsvReader.ReadLine()
	End Function

	Public Function ReadSpecificRow( _
		ByVal intRowNum _
		)

	End Function

	Public Function ReadRows( _
		ByVal intNumOfRows _
		)

		Dim strRows, _
			intIndex

		strRows = ""

		For intIndex = 1 To intNumOfRows
			strRows = strRows & p_CsvReader.ReadLine() & vbCrLf
		Next

		strRows = Left(strRows, Len(strRows) - 1)

		ReadRows = strRows
	End Function

	Public Function ReadSpecificRows( _
		ByVal arrRowNums _
		)

	End Function

	Public Function ReadAll()
		ReadAll = p_CsvReader.ReadAll()
	End Function

	Public Sub Reset()
		With p_CsvReader
			.Close
			.Open p_Filename, FOR_READING
		End With
	End Sub

	Public Sub SkipRow()
		p_CsvReader.SkipLine 
	End Sub

	Public Sub SkipRows( _
		ByVal intNumOfRows _
		)
		
		Dim intIndex

		For intIndex = 1 To intNumOfRows
			p_CsvReader.SkipLine
		Next 
	End Sub

	Private Sub Class_Terminate()
		Set p_CsvReader = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_CSV_Reader.vbs" Then

End If