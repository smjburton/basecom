Option Explicit

Class base_Text_Utf8Encoding
	Private p_Utf8Encoding

	Private Sub Class_Initialize()
		Set p_Utf8Encoding = CreateObject("System.Text.UTF8Encoding")
	End Sub


	' Properties


	Property Get BodyName()
		BodyName = p_Utf8Encoding.BodyName
	End Property

	Property Get CodePage()
		CodePage = p_Utf8Encoding.CodePage
	End Property

	Property Get DecoderFallback()
		Set DecoderFallback = p_Utf8Encoding.DecoderFallback
	End Property

	Property Set DecoderFallback(objDecoderFallback)
		Set p_Utf8Encoding.DecoderFallback = objDecoderFallback
	End Property

	Property Get EncoderFallback()
		Set EncoderFallback = p_Utf8Encoding.EncoderFallback
	End Property

	Property Set EncoderFallback(objEncoderFallback)
		Set p_Utf8Encoding.EncoderFallback = objEncoderFallback
	End Property

	Property Get EncodingName()
		EncodingName = p_Utf8Encoding.EncodingName
	End Property

	Property Get HeaderName()
		HeaderName = p_Utf8Encoding.HeaderName
	End Property

	Property Get IsBrowserDisplay()
		IsBrowserDisplay = p_Utf8Encoding.IsBrowserDisplay
	End Property

	Property Get IsBrowserSave()
		IsBrowserSave = p_Utf8Encoding.IsBrowserSave
	End Property

	Property Get IsMailNewsDisplay()
		IsMailNewsDisplay = p_Utf8Encoding.IsMailNewsDisplay
	End Property

	Property Get IsMailNewsSave()
		IsMailNewsSave = p_Utf8Encoding.IsMailNewsSave
	End Property

	Property Get IsReadOnly()
		IsReadOnly = p_Utf8Encoding.IsReadOnly
	End Property

	Property Get IsSingleByte()
		IsSingleByte = p_Utf8Encoding.IsSingleByte
	End Property

	Property Get WebName()
		WebName = p_Utf8Encoding.WebName
	End Property

	Property Get WindowsCodePage()
		WindowsCodePage = p_Utf8Encoding.WindowsCodePage
	End Property


	' Methods


	Public Function Clone()
		Set Clone = p_Utf8Encoding.Clone
	End Function

	Public Function Equals(objObject)
		Equals = p_Utf8Encoding.Equals(objObject)
	End Function

	Public Sub Finalize()
		p_Utf8Encoding.Finalize
	End Sub

	Public Function GetByteCount(ByRef strChar, ByVal intCount)
	' Function GetByteCount(Char()) As Integer
	' Function GetByteCount(Char(), Int32, Int32) As Integer
	' Function GetByteCount(String) As Integer
		GetByteCount = p_Utf8Encoding.GetByteCount(strChar, intCount)
	End Function

	Public Function GetBytes(ByRef strChar, ByVal intCharCount, ByRef bytByte, ByVal intByteCount)
	' Function GetBytes(Char()) As Integer
	' Function GetBytes(Char(), Int32, Int32) As Integer
	' Function GetBytes(Char(), Int32, Int32, Byte(), Int32) As Integer
	' Function GetBytes(String) As Integer
	' Function GetBytes(String, Int32, Int32, Byte(), Int32) As Integer
		GetBytes = p_Utf8Encoding.GetBytes(strChar, intCharCount, bytByte, intByteCount)
	End Function

	Public Function GetCharCount(ByRef bytBytes, ByVal intCount)
	' Function GetCharCount(Byte()) As Integer
	' Function GetCharCount(Byte(), Int32, Int32) As Integer
		GetCharCount = p_Utf8Encoding.GetCharCount(bytBytes, intCount)
	End Function

	Public Function GetChars(ByRef bytBytes, ByVal intByteCount, ByRef strChars, ByVal intCharCount)
	' Function GetChars(Byte()) As Integer
	' Function GetChars(Byte(), Int32, Int32) As Integer
	' Function GetChars(Byte(), Int32, Int32, Char(), Int32) As Integer
		GetChars = p_Utf8Encoding.GetChars(bytBytes, intByteCount, strChars, intCharCount)
	End Function

	Public Function GetDecoder()
		Set GetDecoder = p_Utf8Encoding.GetDecoder()
	End Function

	Public Function GetEncoder()
		Set GetEncoder = p_Utf8Encoding.GetEncoder()
	End Function

	Public Function GetHashCode()
		GetHashCode = p_Utf8Encoding.GetHashCode
	End Function

	Public Function GetMaxByteCount(intCharCount)
		GetMaxByteCount = p_Utf8Encoding.GetMaxByteCount(intCharCount)
	End Function

	Public Function GetMaxCharCount(intByteCount)
		 GetMaxCharCount = p_Utf8Encoding.GetMaxCharCount(intByteCount)
	End Function

	Public Function GetPreamble()
		GetPreamble = p_Utf8Encoding.GetPreamble()
	End Function

	Public Function GetString(arrBytes(), intByteIndex, intByteCount)
	' Function GetString(Byte())As String
	' Function GetString(Byte*, Int32) As String
		GetString = p_Utf8Encoding.GetString(arrBytes(), intByteIndex, intByteCount)
	End Function

	Public Function GetType()
		Set GetType = p_Utf8Encoding.GetType()
	End Function

	Public Function IsAlwaysNormalized()
	' Function IsAlwaysNormalized(NormalizationForm) As Boolean
		IsAlwaysNormalized = p_Utf8Encoding.IsAlwaysNormalized()
	End Function

	Public Function MemberwiseClone()
		Set MemberwiseClone = p_Utf8Encoding.MemberwiseClone()
	End Function

	Public Function ToString()
		ToString = p_Utf8Encoding.ToString()
	End Function

	Private Sub Class_Terminate()
		Set p_Utf8Encoding = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Text_Utf8Encoding.vbs" Then

End If
