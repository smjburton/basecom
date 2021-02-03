Option Explicit

Class base_IO_Stream
	Private p_Stream

	Private Sub Class_Initialize()
		Set p_Stream = CreateObject("ADODB.Stream")
	End Sub


	' Properties


	Public Property Get Charset()
		Charset = p_Stream.Charset 
	End Property

	Public Property Let Charset(strCharset)
		p_Stream.Charset = strCharset
	End Property

	Public Property Get IsEndOfStream()
		IsEndOfStream = p_Stream.EOS
	End Property

	Public Property Get LineSeparator()
		Set LineSeparator = p_Stream.LineSeparator
	End Property

	Public Property Set LineSeparator(objLineSeparatorEnum)
		Set p_Stream.LineSeparator = objLineSeparatorEnum
	End Property

	Public Property Get Mode()
		Set Mode = p_Stream.Mode 
	End Property

	Public Property Set Mode(objConnectModeEnum)
		Set p_Stream.Mode = objConnectModeEnum
	End Property

	Public Property Get Position()
		Position = p_Stream.Position
	End Property

	Public Property Let Position(lngPosition)
		p_Stream.Position = lngPosition
	End Property

	Public Property Get Size()
		Size = p_Stream.Size 
	End Property

	Public Property Get State()
		Set State = p_Stream.State 
	End Property

	Public Property Get StreamType()
		StreamType = p_Stream.Type 
	End Property

	Public Property Let StreamType(intStreamType)
		p_Stream.Type = intStreamType
	End Property


	' Methods


	Public Sub Cancel()
		p_Stream.Cancel
	End Sub

	Public Sub Close()
		p_Stream.Close
	End Sub

	Public Sub CopyTo(objDestStream) ' Optional params: [CharNumber As Long = -1]
		p_Stream.CopyTo objDestStream
	End Sub

	Public Sub Flush()
		p_Stream.Flush
	End Sub

	Public Sub LoadFromFile(strFileName)
		p_Stream.LoadFromFile strFileName
	End Sub

	Public Sub Open() ' Optional params: [Source], [Mode As ConnectModeEnum = adModeUnknown], [Options As StreamOpenOptionsEnum = adOpenStreamUnspecified], [UserName As String], [Password As String])
		p_Stream.Open
	End Sub

	Public Function Read() ' Optional params: [NumBytes As Long = -1]
		p_Stream.Read
	End Function

	Public Function ReadText() ' Optional params: [NumChars As Long = -1]) As String
		ReadText = p_Stream.ReadText()
	End Function

	Public Sub SaveToFile(strFileName) ' Optional params: [Options As SaveOptionsEnum = adSaveCreateNotExist])
		p_Stream.SaveToFile strFileName
	End Sub

	Public Sub SetEOS()
		p_Stream.SetEOS
	End Sub

	Public Sub SkipLine()
		p_Stream.SkipLine
	End Sub

	Public Sub Write(arrBuffer())
		p_Stream.Write arrBuffer
	End Sub

	Public Sub WriteText(strData) ' Optional params: [Options As StreamWriteEnum = adWriteChar])
		p_Stream.WriteText strData
	End Sub

	Private Sub Class_Terminate()
		Set p_Stream = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_IO_Stream.vbs" Then

End If
