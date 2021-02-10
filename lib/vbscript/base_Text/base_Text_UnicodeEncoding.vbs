Option Explicit

Class base_Text_UnicodeEncoding
	Private p_UnicodeEncoding

	Private Sub Class_Initialize()
		Set p_UnicodeEncoding = CreateObject("System.Text.UnicodeEncoding")
	End Sub


	' Properties


	Public Property Get BodyName()
		BodyName = p_UnicodeEncoding.BodyName
	End Property

	Public Property Get CodePage()
		CodePage = p_UnicodeEncoding.CodePage	
	End Property

	Public Property Get DecoderFallback()
		Set DecoderFallback = p_UnicodeEncoding.DecoderFallback
	End Property

	Public Property Set DecoderFallback(objDecoderFallback)
		Set DecoderFallback = p_UnicodeEncoding.DecoderFallback(objDecoderFallback)
	End Property

	Public Property Get EncoderFallback()
		Set EncoderFallback = p_UnicodeEncoding.EncoderFallback
	End Property

	Public Property Set EncoderFallback(objEncoderFallback)
		Set EncoderFallback = p_UnicodeEncoding.EncoderFallback(objEncoderFallback)
	End Property

	Public Property Get EncodingName()
		EncodingName = p_UnicodeEncoding.EncodingName
	End Property

	Public Property Get HeaderName()
		HeaderName = p_UnicodeEncoding.HeaderName	
	End Property

	Public Property Get IsBrowserDisplay()
		IsBrowserDisplay = p_UnicodeEncoding.IsBrowserDisplay	
	End Property

	Public Property Get IsBrowserSave()
		IsBrowserSave = p_UnicodeEncoding.IsBrowserSave
	End Property

	Public Property Get IsMailNewsDisplay()
		IsMailNewsDisplay = p_UnicodeEncoding.IsMailNewsDisplay	
	End Property

	Public Property Get IsMailNewsSave()
		IsMailNewsSave = p_UnicodeEncoding.IsMailNewsSave
	End Property

	Public Property Get IsReadOnly()
		IsReadOnly = p_UnicodeEncoding.IsReadOnly	
	End Property

	Public Property Get IsSingleByte()
		IsSingleByte = p_UnicodeEncoding.IsSingleByte	
	End Property

	Public Property Get WebName()
		WebName = p_UnicodeEncoding.WebName	
	End Property

	Public Property Get WindowsCodePage()
		WindowsCodePage = p_UnicodeEncoding.WindowsCodePage	
	End Property


	' Methods


	Public Function Clone()
		Set Clone = p_UnicodeEncoding.Clone()
	End Function

	Public Function Equals(objObject)
		Equals = p_UnicodeEncoding.Equals(objObject)
	End Function

	Public Sub Finalize()
		p_UnicodeEncoding.Finalize	
	End Sub

	Public Function GetByteCount(ByRef strChar, ByVal intCount)
	' Function GetByteCount(Char()) As Integer
	' Function GetByteCount(Char(), Int32, Int32) As Integer
	' Function GetByteCount(String) As Integer
		GetByteCount = p_UnicodeEncoding.GetByteCount(strChar, intCount)
	End Function

	Public Function GetBytes(ByRef strChar, ByVal intCharCount, ByRef bytByte, ByVal intByteCount)
	' Function GetBytes(Char()) As Integer
	' Function GetBytes(Char(), Int32, Int32) As Integer
	' Function GetBytes(Char(), Int32, Int32, Byte(), Int32) As Integer
	' Function GetBytes(String) As Integer
	' Function GetBytes(String, Int32, Int32, Byte(), Int32) As Integer
		GetBytes = p_UnicodeEncoding.GetBytes(strChar, intCharCount, bytByte, intByteCount)
	End Function

	Public Function GetCharCount(ByRef bytBytes, ByVal intCount)
	' Function GetCharCount(Byte()) As Integer
	' Function GetCharCount(Byte(), Int32, Int32) As Integer
		GetCharCount = p_UnicodeEncoding.GetCharCount(bytBytes, intCount)
	End Function

	Public Function GetChars(ByRef bytBytes, ByVal intByteCount, ByRef strChars, ByVal intCharCount)
	' Function GetChars(Byte()) As Integer
	' Function GetChars(Byte(), Int32, Int32) As Integer
	' Function GetChars(Byte(), Int32, Int32, Char(), Int32) As Integer
		GetChars = p_UnicodeEncoding.GetChars(bytBytes, intByteCount, strChars, intCharCount)
	End Function

	Public Function GetDecoder()
		Set GetDecoder = p_UnicodeEncoding.GetDecoder()
	End Function

	Public Function GetEncoder()
		Set GetEncoder = p_UnicodeEncoding.GetEncoder()
	End Function

	Public Function GetHashCode()
		GetHashCode = p_UnicodeEncoding.GetHashCode
	End Function

	Public Function GetMaxByteCount(intCharCount)
		GetMaxByteCount = p_UnicodeEncoding.GetMaxByteCount(intCharCount)
	End Function

	Public Function GetMaxCharCount(intByteCount)
		GetMaxCharCount = p_UnicodeEncoding.GetMaxCharCount(intByteCount)
	End Function

	Public Function GetPreamble()
		GetPreamble = p_UnicodeEncoding.GetPreamble()
	End Function

	Public Function GetString(arrBytes(), intByteIndex, intByteCount)
	' Function GetString(Byte()) As String
	' Function GetString(Byte*, Int32) As String
		GetString = p_UnicodeEncoding.GetString(arrBytes(), intByteIndex, intByteCount)
	End Function

	Public Function GetType()
		Set GetType = p_UnicodeEncoding.GetType()
	End Function

	Public Function IsAlwaysNormalized()
	' Function IsAlwaysNormalized(NormalizationForm) As Boolean
		IsAlwaysNormalized = p_UnicodeEncoding.IsAlwaysNormalized()
	End Function

	Public Function MemberwiseClone()
		Set MemberwiseClone = p_UnicodeEncoding.MemberwiseClone()
	End Function

	Public Function ToString()
		ToString = p_UnicodeEncoding.ToString()
	End Function

	Private Sub Class_Terminate()
		Set p_UnicodeEncoding = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Text_UnicodeEncoding.vbs" Then

End If
