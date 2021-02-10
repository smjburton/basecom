Option Explicit

Class base_Text_AsciiEncoding
	Private p_AsciiEncoding

	Private Sub Class_Initialize()
		Set p_AsciiEncoding = CreateObject("System.Text.ASCIIEncoding")
	End Sub


	' Properties


	Property Get BodyName()
		BodyName = p_AsciiEncoding.BodyName
	End Property

	Property Get CodePage()
		CodePage = p_AsciiEncoding.CodePage
	End Property

	Property Get DecoderFallback()
		Set DecoderFallback = p_AsciiEncoding.DecoderFallback
	End Property

	Property Set DecoderFallback(objDecoderFallback)
		Set p_AsciiEncoding.DecoderFallback = objDecoderFallback
	End Property

	Property Get EncoderFallback()
		Set EncoderFallback = p_AsciiEncoding.EncoderFallback
	End Property

	Property Set EncoderFallback(objEncoderFallback)
		Set p_AsciiEncoding.EncoderFallback = objEncoderFallback
	End Property

	Property Get EncodingName()
		EncodingName = p_AsciiEncoding.EncodingName
	End Property

	Property Get HeaderName()
		HeaderName = p_AsciiEncoding.HeaderName
	End Property

	Property Get IsBrowserDisplay()
		IsBrowserDisplay = p_AsciiEncoding.IsBrowserDisplay
	End Property

	Property Get IsBrowserSave()
		IsBrowserSave = p_AsciiEncoding.IsBrowserSave
	End Property

	Property Get IsMailNewsDisplay()
		IsMailNewsDisplay = p_AsciiEncoding.IsMailNewsDisplay
	End Property

	Property Get IsMailNewsSave()
		IsMailNewsSave = p_AsciiEncoding.IsMailNewsSave
	End Property

	Property Get IsReadOnly()
		IsReadOnly = p_AsciiEncoding.IsReadOnly
	End Property

	Property Get IsSingleByte()
		IsSingleByte = p_AsciiEncoding.IsSingleByte
	End Property

	Property Get WebName()
		WebName = p_AsciiEncoding.WebName
	End Property

	Property Get WindowsCodePage()
		WindowsCodePage = p_AsciiEncoding.WindowsCodePage
	End Property


	' Methods


	Public Function Clone()
		Set Clone = p_AsciiEncoding.Clone
	End Function

	Public Function Equals(objObject)
		Equals = p_AsciiEncoding.Equals(objObject)
	End Function

	Public Sub Finalize()
		p_AsciiEncoding.Finalize
	End Sub

	Public Function GetByteCount(ByRef strChar, ByVal intCount)
	' Function GetByteCount(Char()) As Integer
	' Function GetByteCount(Char(), Int32, Int32) As Integer
	' Function GetByteCount(String) As Integer
		GetByteCount = p_AsciiEncoding.GetByteCount(strChar, intCount)
	End Function

	Public Function GetBytes(ByRef strChar, ByVal intCharCount, ByRef bytByte, ByVal intByteCount)
	' Function GetBytes(Char()) As Integer
	' Function GetBytes(Char(), Int32, Int32) As Integer
	' Function GetBytes(Char(), Int32, Int32, Byte(), Int32) As Integer
	' Function GetBytes(String) As Integer
	' Function GetBytes(String, Int32, Int32, Byte(), Int32) As Integer
		GetBytes = p_AsciiEncoding.GetBytes(strChar, intCharCount, bytByte, intByteCount)
	End Function

	Public Function GetCharCount(ByRef bytBytes, ByVal intCount)
	' Function GetCharCount(Byte()) As Integer
	' Function GetCharCount(Byte(), Int32, Int32) As Integer
		GetCharCount = p_AsciiEncoding.GetCharCount(bytBytes, intCount)
	End Function

	Public Function GetChars(ByRef bytBytes, ByVal intByteCount, ByRef strChars, ByVal intCharCount)
	' Function GetChars(Byte()) As Integer
	' Function GetChars(Byte(), Int32, Int32) As Integer
	' Function GetChars(Byte(), Int32, Int32, Char(), Int32) As Integer
		GetChars = p_AsciiEncoding.GetChars(bytBytes, intByteCount, strChars, intCharCount)
	End Function

	Public Function GetDecoder()
		Set GetDecoder = p_AsciiEncoding.GetDecoder()
	End Function

	Public Function GetEncoder()
		Set GetEncoder = p_AsciiEncoding.GetEncoder()
	End Function

	Public Function GetHashCode()
		GetHashCode = p_AsciiEncoding.GetHashCode
	End Function

	Public Function GetMaxByteCount(intCharCount)
		GetMaxByteCount = p_AsciiEncoding.GetMaxByteCount(intCharCount)
	End Function

	Public Function GetMaxCharCount(intByteCount)
		 GetMaxCharCount = p_AsciiEncoding.GetMaxCharCount(intByteCount)
	End Function

	Public Function GetPreamble()
		GetPreamble = p_AsciiEncoding.GetPreamble()
	End Function

	Public Function GetString(arrBytes(), intByteIndex, intByteCount)
	' Function GetString(Byte())As String
	' Function GetString(Byte*, Int32) As String
		GetString = p_AsciiEncoding.GetString(arrBytes(), intByteIndex, intByteCount)
	End Function

	Public Function GetType()
		Set GetType = p_AsciiEncoding.GetType()
	End Function

	Public Function IsAlwaysNormalized()
	' Function IsAlwaysNormalized(NormalizationForm) As Boolean
		IsAlwaysNormalized = p_AsciiEncoding.IsAlwaysNormalized()
	End Function

	Public Function MemberwiseClone()
		Set MemberwiseClone = p_AsciiEncoding.MemberwiseClone()
	End Function

	Public Function ToString()
		ToString = p_AsciiEncoding.ToString()
	End Function

	Private Sub Class_Terminate()
		Set p_AsciiEncoding = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Text_AsciiEncoding.vbs" Then

End If
