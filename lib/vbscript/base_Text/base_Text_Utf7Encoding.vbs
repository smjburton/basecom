Option Explicit

Class base_Text_Utf7Encoding
	Private p_Utf7Encoding

	Private Sub Class_Initialize()
		Set p_Utf7Encoding = CreateObject("System.Text.UTF7Encoding")
	End Sub

	
	' Properties


	Property Get BodyName()
		BodyName = p_Utf7Encoding.BodyName
	End Property

	Property Get CodePage()
		CodePage = p_Utf7Encoding.CodePage
	End Property

	Property Get DecoderFallback()
		Set DecoderFallback = p_Utf7Encoding.DecoderFallback
	End Property

	Property Set DecoderFallback(objDecoderFallback)
		Set p_Utf7Encoding.DecoderFallback = objDecoderFallback
	End Property

	Property Get EncoderFallback()
		Set EncoderFallback = p_Utf7Encoding.EncoderFallback
	End Property

	Property Set EncoderFallback(objEncoderFallback)
		Set p_Utf7Encoding.EncoderFallback = objEncoderFallback
	End Property

	Property Get EncodingName()
		EncodingName = p_Utf7Encoding.EncodingName
	End Property

	Property Get HeaderName()
		HeaderName = p_Utf7Encoding.HeaderName
	End Property

	Property Get IsBrowserDisplay()
		IsBrowserDisplay = p_Utf7Encoding.IsBrowserDisplay
	End Property

	Property Get IsBrowserSave()
		IsBrowserSave = p_Utf7Encoding.IsBrowserSave
	End Property

	Property Get IsMailNewsDisplay()
		IsMailNewsDisplay = p_Utf7Encoding.IsMailNewsDisplay
	End Property

	Property Get IsMailNewsSave()
		IsMailNewsSave = p_Utf7Encoding.IsMailNewsSave
	End Property

	Property Get IsReadOnly()
		IsReadOnly = p_Utf7Encoding.IsReadOnly
	End Property

	Property Get IsSingleByte()
		IsSingleByte = p_Utf7Encoding.IsSingleByte
	End Property

	Property Get WebName()
		WebName = p_Utf7Encoding.WebName
	End Property

	Property Get WindowsCodePage()
		WindowsCodePage = p_Utf7Encoding.WindowsCodePage
	End Property


	' Methods


	Public Function Clone()
		Set Clone = p_Utf7Encoding.Clone
	End Function

	Public Function Equals(objObject)
		Equals = p_Utf7Encoding.Equals(objObject)
	End Function

	Public Sub Finalize()
		p_Utf7Encoding.Finalize
	End Sub

	Public Function GetByteCount(ByRef strChar, ByVal intCount)
	' Function GetByteCount(Char()) As Integer
	' Function GetByteCount(Char(), Int32, Int32) As Integer
	' Function GetByteCount(String) As Integer
		GetByteCount = p_Utf7Encoding.GetByteCount(strChar, intCount)
	End Function

	Public Function GetBytes(ByRef strChar, ByVal intCharCount, ByRef bytByte, ByVal intByteCount)
	' Function GetBytes(Char()) As Integer
	' Function GetBytes(Char(), Int32, Int32) As Integer
	' Function GetBytes(Char(), Int32, Int32, Byte(), Int32) As Integer
	' Function GetBytes(String) As Integer
	' Function GetBytes(String, Int32, Int32, Byte(), Int32) As Integer
		GetBytes = p_Utf7Encoding.GetBytes(strChar, intCharCount, bytByte, intByteCount)
	End Function

	Public Function GetCharCount(ByRef bytBytes, ByVal intCount)
	' Function GetCharCount(Byte()) As Integer
	' Function GetCharCount(Byte(), Int32, Int32) As Integer
		GetCharCount = p_Utf7Encoding.GetCharCount(bytBytes, intCount)
	End Function

	Public Function GetChars(ByRef bytBytes, ByVal intByteCount, ByRef strChars, ByVal intCharCount)
	' Function GetChars(Byte()) As Integer
	' Function GetChars(Byte(), Int32, Int32) As Integer
	' Function GetChars(Byte(), Int32, Int32, Char(), Int32) As Integer
		GetChars = p_Utf7Encoding.GetChars(bytBytes, intByteCount, strChars, intCharCount)
	End Function

	Public Function GetDecoder()
		Set GetDecoder = p_Utf7Encoding.GetDecoder()
	End Function

	Public Function GetEncoder()
		Set GetEncoder = p_Utf7Encoding.GetEncoder()
	End Function

	Public Function GetHashCode()
		GetHashCode = p_Utf7Encoding.GetHashCode
	End Function

	Public Function GetMaxByteCount(intCharCount)
		GetMaxByteCount = p_Utf7Encoding.GetMaxByteCount(intCharCount)
	End Function

	Public Function GetMaxCharCount(intByteCount)
		 GetMaxCharCount = p_Utf7Encoding.GetMaxCharCount(intByteCount)
	End Function

	Public Function GetPreamble()
		GetPreamble = p_Utf7Encoding.GetPreamble()
	End Function

	Public Function GetString(arrBytes(), intByteIndex, intByteCount)
	' Function GetString(Byte())As String
	' Function GetString(Byte*, Int32) As String
		GetString = p_Utf7Encoding.GetString(arrBytes(), intByteIndex, intByteCount)
	End Function

	Public Function GetType()
		Set GetType = p_Utf7Encoding.GetType()
	End Function

	Public Function IsAlwaysNormalized()
	' Function IsAlwaysNormalized(NormalizationForm) As Boolean
		IsAlwaysNormalized = p_Utf7Encoding.IsAlwaysNormalized()
	End Function

	Public Function MemberwiseClone()
		Set MemberwiseClone = p_Utf7Encoding.MemberwiseClone()
	End Function

	Public Function ToString()
		ToString = p_Utf7Encoding.ToString()
	End Function

	Private Sub Class_Terminate()
		Set p_Utf7Encoding = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Text_Utf7Encoding.vbs" Then

End If
