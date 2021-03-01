Option Explicit

Class base_Database_DAO_Field
	Private p_Field

	Private Sub Class_Initialize()
		Set p_Field = CreateObject("DAO.Field.36")
	End Sub


	' Properties


	Public Property Get AllowZeroLength()
		AllowZeroLength = p_Field.AllowZeroLength
	End Property

	Public Property Let AllowZeroLength(blnAllowZeroLength)
		p_Field.AllowZeroLength = blnAllowZeroLength
	End Property

	Public Property Get Attributes()
		Attributes = p_Field.Attributes
	End Property

	Public Property Let Attributes(lngAttributes)
		p_Field.Attributes = lngAttributes
	End Property

	Public Property Get CollatingOrder()
		CollatingOrder = p_Field.CollatingOrder
	End Property

	Public Property Get DataUpdatable()
		DataUpdatable = p_Field.DataUpdatable
	End Property

	Public Property Get DefaultValue()
		DefaultValue = p_Field.DefaultValue
	End Property

	Public Property Let DefaultValue(varDefaultValue)
		p_Field.DefaultValue = varDefaultValue
	End Property

	Public Property Set DefaultValue(varDefaultValue)
		Set p_Field.DefaultValue = varDefaultValue
	End Property

	Public Property Get FieldSize()
		FieldSize = p_Field.FieldSize
	End Property

	Public Property Get ForeignName()
		ForeignName = p_Field.ForeignName
	End Property

	Public Property Let ForeignName(strForeignName)
		p_Field.ForeignName = strForeignName
	End Property

	Public Property Get Name()
		Name = p_Field.Name
	End Property

	Public Property Let Name(strName)
		p_Field.Name = strName
	End Property

	Public Property Get OrdinalPosition()
		OrdinalPosition = p_Field.OrdinalPosition
	End Property

	Public Property Let OrdinalPosition(intOrdinalPosition)
		p_Field.OrdinalPosition = intOrdinalPosition
	End Property

	Public Property Get OriginalValue()
		OriginalValue = p_Field.OriginalValue
	End Property

	Public Property Get Properties()
		Set Properties = p_Field.Properties
	End Property

	Public Property Get Required()
		Required = p_Field.Required
	End Property

	Public Property Let Required(blnRequired)
		p_Field.Required = blnRequired
	End Property

	Public Property Get Size()
		Size = p_Field.Size
	End Property

	Public Property Let Size(lngSize)
		p_Field.Size = lngSize
	End Property

	Public Property Get SourceField()
		SourceField = p_Field.SourceField
	End Property

	Public Property Get SourceTable(strSourceTable)
		p_Field.SourceTable = strSourceTable
	End Property

	Public Property Get FieldType()
		FieldType = p_Field.Type
	End Property

	Public Property Let FieldType(intType)
		p_Field.Type = intType
	End Property

	Public Property Get ValidateOnSet()
		ValidateOnSet = p_Field.ValidateOnSet
	End Property

	Public Property Let ValidateOnSet(blnValidateOnSet)
		p_Field.ValidateOnSet = blnValidateOnSet
	End Property

	Public Property Get ValidationRule()
		ValidationRule = p_Field.ValidationRule
	End Property

	Public Property Let ValidationRule(strValidationRule)
		p_Field.ValidationRule = strValidationRule
	End Property

	Public Property Get ValidationText()
		ValidationText = p_Field.ValidationText
	End Property

	Public Property Let ValidationText(strValidationText)
		p_Field.ValidationText = strValidationText
	End Property

	Public Default Property Get Value()
		If IsObject(p_Field.Value) Then
			Set Value = p_Field.Value
		Else
			Value = p_Field.Value
		End If
	End Property

	Public Property Let Value(varValue)
		p_Field.Value = strValue
	End Property

	Public Property Set Value(varValue)
		Set Value = p_Field.Value
	End Property

	Public Property Get VisibleValue()
		VisibleValue = p_Field.VisibleValue
	End Property


	' Methods


	Public Sub AppendChunk(varValue)
		p_Field.AppendChunk varValue
	End Sub

	Public Function CreateProperty() ' Optional parameters: [Name], [Type], [Value], [DDL]
		Set CreateProperty = p_Field.CreateProperty()
	End Function

	Public Function GetChunk(lngOffset, lngBytes)
		If IsObject(p_Field.GetChunk(lngOffset, lngBytes)) Then
			Set GetChunk = p_Field.GetChunk(lngOffset, lngBytes)
		Else
			GetChunk = p_Field.GetChunk(lngOffset, lngBytes)
		End If
	End Function

	Private Sub Class_Terminate()
		Set p_Field = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_DAO_Field.vbs" Then

End If
