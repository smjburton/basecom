Option Explicit

Class base_IO_StringWriter
	Private p_StringWriter
	
	Private Sub Class_Initialize()
		Set p_StringWriter = CreateObject("System.IO.StringWriter")
	End Sub


	' Properties


	Public Property Get Encoding()
		Set Encoding = p_StringWriter.Encoding 
	End Property

	Public Property Get FormatProvider()
		Set FormatProvider = p_StringWriter.FormatProvider
	End Property

	Public Property Get NewLine()
		NewLine = p_StringWriter.NewLine 
	End Property

	Public Property Let NewLine(strNewLine)
		p_StringWriter.NewLine = strNewLine
	End Property


	' Methods


	Public Sub Close()
		p_StringWriter.Close
	End Sub

	Public Function CreateObjRef(objType)
		Set CreateObjRef = p_StringWriter.CreateObjRef(objType)
	End Function

	Public Sub Dispose()
	' Sub Dispose(Boolean)
		p_StringWriter.Dispose
	End Sub

	Public Function Equals(objObject)
		Equals = p_StringWriter.Equals(objObject)
	End Function

	Public Sub Finalize()
		p_StringWriter.Finalize
	End Sub

	Public Sub Flush()
		p_StringWriter.Flush
	End Sub

	Public Function FlushAsync()
		Set FlushAsync = p_StringWriter.FlushAsync()
	End Function

	Public Function GetHashCode()
		GetHashCode = p_StringWriter.GetHashCode()
	End Function

	Public Function GetLifetimeService()
		Set GetLifetimeService = p_StringWriter.GetLifetimeService()
	End Function

	Public Function GetStringBuilder()
		Set GetStringBuilder = p_StringWriter.GetStringBuilder()
	End Function

	Public Function GetType()
		Set GetType = p_StringWriter.GetType()
	End Function

	Public Function InitializeLifetimeService()
		Set InitializeLifetimeService = p_StringWriter.InitializeLifetimeService()
	End Function

	Public Function MemberwiseClone()
	' Function MemberwiseClone(Boolean) As Object
		Set MemberwiseClone = p_StringWriter.MemberwiseClone()
	End Function

	Public Function ToString()
		ToString = p_StringWriter.ToString()
	End Function

	Public Sub Write(blnBoolean)
	' Sub Write(Char)
	' Sub Write(Char())
	' Sub Write(Char(), Int32, Int32)
	' Sub Write(Decimal)
	' Sub Write(Double)
	' Sub Write(Int32)
	' Sub Write(Int64)
	' Sub Write(Object)
	' Sub Write(Single)
	' Sub Write(String)
	' Sub Write(String, Object)
	' Sub Write(String, Object, Object)
	' Sub Write(String, Object, Object, Object)
	' Sub Write(String, Object())
	' Sub Write(UInt32)
	' Sub Write(UInt64)
		p_StringWriter.Write blnBoolean
	End Sub

	Public Function WriteAsync(strValue)
	' Function WriteAsync(Char())As System.Threading.Tasks.Task
	' Function WriteAsync(Char(), Int32, Int32) As System.Threading.Tasks.Task
	' Function WriteAsync(String) As System.Threading.Tasks.Task
		Set WriteAsync = p_StringWriter.WriteAsync(strValue)
	End Function

	Public Sub WriteLine()
	' Sub WriteLine(Boolean)
	' Sub WriteLine(Char)
	' Sub WriteLine(Char())
	' Sub WriteLine(Char(), Int32, Int32)
	' Sub WriteLine(Decimal)
	' Sub WriteLine(Double)
	' Sub WriteLine(Int32)
	' Sub WriteLine(Int64)
	' Sub WriteLine(Object)
	' Sub WriteLine(Single)
	' Sub WriteLine(String)
	' Sub WriteLine(String, Object)
	' Sub WriteLine(String, Object, Object)
	' Sub WriteLine(String, Object, Object, Object)
	' Sub WriteLine(String, Object())
	' Sub WriteLine(UInt32)
	' Sub WriteLine(UInt64)
		p_StringWriter.WriteLine
	End Sub

	Public Function WriteLineAsync()
	' Function WriteLineAsync(Char) As System.Threading.Tasks.Task
	' Function WriteLineAsync(Char())As System.Threading.Tasks.Task
	' Function WriteLineAsync(Char(), Int32, Int32) As System.Threading.Tasks.Task
	' Function WriteLineAsync(String) As System.Threading.Tasks.Task
		Set WriteLineAsync = p_StringWriter.WriteLineAsync()
	End Function

	Private Sub Class_Terminate()
		Set p_StringWriter = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_IO_StringWriter.vbs" Then

End If
