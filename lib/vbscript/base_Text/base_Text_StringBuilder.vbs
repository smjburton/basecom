Option Explicit

Class base_Text_StringBuilder
	Private p_StringBuilder

	Private Sub Class_Initialize()
		Set p_StringBuilder = CreateObject("System.Text.StringBuilder")
	End Sub


	' Properties


	Public Property Get Capacity()
		Capacity = p_StringBuilder.Capacity
	End Property

	Public Property Let Capacity(intCapacity)
		p_StringBuilder.Capacity = intCapacity
	End Property

	Public Property Get Chars(intIndex)
		Chars = p_StringBuilder.Chars(intIndex)
	End Property

	Public Property Let Chars(intIndex, strChar)
		p_StringBuilder.Chars(intIndex) = strChar
	End Property

	Public Property Get Length()
		Length = p_StringBuilder.Length
	End Property

	Public Property Let Length(intLength)
		p_StringBuilder.Length = intLength
	End Property

	Public Property Get MaxCapacity()
		MaxCapacity = p_StringBuilder.MaxCapacity
	End Property


	' Methods


	Public Function Append(strBoolean)
	' Function Append(Byte) As StringBuilder
	' Function Append(Char) As StringBuilder
	' Function Append(Char*, Int32) As StringBuilder
	' Function Append(Char, Int32) As StringBuilder
	' Function Append(Char())As StringBuilder
	' Function Append(Char(), Int32, Int32) As StringBuilder
	' Function Append(Decimal) As StringBuilder
	' Function Append(Double) As StringBuilder
	' Function Append(Int16) As StringBuilder
	' Function Append(Int32) As StringBuilder
	' Function Append(Int64) As StringBuilder
	' Function Append(Object) As StringBuilder
	' Function Append(SByte) As StringBuilder
	' Function Append(Single) As StringBuilder
	' Function Append(String) As StringBuilder
	' Function Append(String, Int32, Int32) As StringBuilder
	' Function Append(UInt16) As StringBuilder
	' Function Append(UInt32) As StringBuilder
	' Function Append(UInt64) As StringBuilder
		Set Append = p_StringBuilder.Append(strBoolean)
	End Function

	Public Function AppendFormat(objFormatProvider, strFormat, objObject)
	' Function AppendFormat(IFormatProvider, String, Object, Object) As StringBuilder
	' Function AppendFormat(IFormatProvider, String, Object, Object, Object) As StringBuilder
	' Function AppendFormat(IFormatProvider, String, Object()) As StringBuilder
	' Function AppendFormat(String, Object) As StringBuilder
	' Function AppendFormat(String, Object, Object) As StringBuilder
	' Function AppendFormat(String, Object, Object, Object) As StringBuilder
	' Function AppendFormat(String, Object()) As StringBuilder
		Set AppendFormat = p_StringBuilder.AppendFormat(objFormatProvider, strFormat, objObject)
	End Function

	Public Function AppendLine()
	' Function AppendLine(String) As StringBuilder
		Set AppendLine = p_StringBuilder.AppendLine()
	End Function

	Public Function Clear()
		Set Clear = p_StringBuilder.Clear()
	End Function

	Public Sub CopyTo(intSourceIndex, strDestinationChar(), intDestinationIndex, intCount)
		p_StringBuilder.CopyTo intSourceIndex, strDestinationChar(), intDestinationIndex, intCount
	End Sub

	Public Function EnsureCapacity(intCapacity)
		EnsureCapacity = p_StringBuilder.EnsureCapacity(intCapacity)
	End Function

	Public Function Equals(objObject)
	' Function Equals(StringBuilder) As Boolean
		Equals = p_StringBuilder.Equals(objObject)
	End Function

	Public Function GetHashCode()
		GetHashCode = p_StringBuilder.GetHashCode()
	End Function

	Public Function GetType()
		Set GetType = p_StringBuilder.GetType()
	End Function

	Public Function Insert(intIndex, blnBooleanValue)
	' Function Insert(Int32, Byte) As StringBuilder
	' Function Insert(Int32, Char) As StringBuilder
	' Function Insert(Int32, Char()) As StringBuilder
	' Function Insert(Int32, Char(), Int32, Int32) As StringBuilder
	' Function Insert(Int32, Decimal) As StringBuilder
	' Function Insert(Int32, Double) As StringBuilder
	' Function Insert(Int32, Int16) As StringBuilder
	' Function Insert(Int32, Int32) As StringBuilder
	' Function Insert(Int32, Int64) As StringBuilder
	' Function Insert(Int32, Object) As StringBuilder
	' Function Insert(Int32, SByte) As StringBuilder
	' Function Insert(Int32, Single) As StringBuilder
	' Function Insert(Int32, String) As StringBuilder
	' Function Insert(Int32, String, Int32) As StringBuilder
	' Function Insert(Int32, UInt16) As StringBuilder
	' Function Insert(Int32, UInt32) As StringBuilder
	' Function Insert(Int32, UInt64) As StringBuilder
		Set Insert = p_StringBuilder.Insert(intIndex, blnBooleanValue)
	End Function

	Public Function Remove(intStartIndex, intLength)
		Set Remove = p_StringBuilder.Remove(intStartIndex, intLength)
	End Function

	Public Function Replace(strOldChar, strNewChar)
	' Function Replace(Char, Char, Int32, Int32) As StringBuilder
	' Function Replace(String, String) As StringBuilder
	' Function Replace(String, String, Int32, Int32) As StringBuilder
		Set Replace = p_StringBuilder.Replace(strOldChar, strNewChar)
	End Function

	Public Function ToStr
	' Function ToString(Int32, Int32) As String
		ToString = p_StringBuilder.ToString()
	End Function

	Private Sub Class_Terminate()
		Set p_StringBuilder = CreateObject("System.Text.StringBuilder")
	End Sub
End Class

If WScript.ScriptName = "base_Text_StringBuilder.vbs" Then

End If
