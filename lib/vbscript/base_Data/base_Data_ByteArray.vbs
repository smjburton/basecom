Option Explicit

Class base_Data_ByteArray
	Private p_objByteArray

	Private Sub Class_Initialize()
		Set p_objByteArray = CreateObject("MSXML2.DOMDocument").CreateElement("Binary")
		p_objByteArray.DataType = "Bin.Hex"
	End Sub

	Public Default Property Get Value()
		Value = p_objByteArray.NodeTypedValue
	End Property

	Public Property Let Value( _
		strValue _
		)

		p_objByteArray.Text = strValue
	End Property

	Private Sub Class_Terminate()
		Set p_objByteArray = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Data_ByteArray.vbs" Then

End If
