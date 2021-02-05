Option Explicit

Class base_Database_Parameter
	Private p_Parameter

	Private Sub Class_Initialize()
		Set p_Parameter = CreateObject("ADODB.Parameter")
	End Sub

	
	' Properties


	Public Property Get Attributes()
		Attributes = p_Parameter.Attributes
	End Property

	Public Property Let Attributes(lngAttributes)
		p_Parameter.Attributes = lngAttributes
	End Property

	Public Property Get Direction()
		Direction = p_Parameter.Direction 
	End Property

	Public Property Let Direction(intParameterDirectionEnum)
		p_Parameter.Direction = intParameterDirectionEnum
	End Property

	Public Property Get Name()
		Name = p_Parameter.Name
	End Property

	Public Property Let Name(strName)
		p_Parameter.Name = strName
	End Property

	Public Property Get NumericScale()
		NumericScale = p_Parameter.NumericScale 
	End Property

	Public Property Let NumericScale(bytNumbericScale)
		p_Parameter.NumericScale = bytNumbericScale
	End Property

	Public Property Get Precision()
		Precision = p_Parameter.Precision
	End Property

	Public Property Let Precision(bytPrecision)
		p_Parameter.Precision = bytPrecision
	End Property

	Public Property Get Properties
		Set Properties = p_Parameter.Properties
	End Property

	Public Property Get Size()
		Size = p_Parameter.Size
	End Property

	Public Property Let Size(lngSize)
		p_Parameter.Size = lngSize
	End Property

	Public Property Get ParamType()
		ParamType = p_Parameter.Type
	End Property

	Public Property Let ParamType(intDataTypeEnum)
		p_Parameter.Type = intDataTypeEnum
	End Property

	Public Property Get Value()
		If IsObject(p_Parameter.Value) Then
			Set Value = p_Parameter.Value
		Else
			Value = p_Parameter.Value
		End If
	End Property

	Public Property Let Value(varValue)
		p_Parameter.Value = varValue
	End Property

	Public Property Set Value(varValue)
		Set p_Parameter.Value = varValue
	End Property


	' Methods


	Public Sub AppendChunk(varVal) 
		p_Parameter.AppendChunk varVal
	End Sub

	Private Sub Class_Terminate()
		Set p_Parameter = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_Paramter.vbs" Then

End If
