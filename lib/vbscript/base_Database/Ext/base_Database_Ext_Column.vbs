Option Explicit

Class base_Database_Ext_Column
	Private p_Column

	Private Sub Class_Initialize()
		Set p_Column = CreateObject("ADOX.Column")
	End Sub


	' Properties


	Public Property Get Attributes()
		Attributes = p_Column.Attributes
	End Property

	Public Property Let Attributes(intColumnAttributesEnum)
		p_Column.Attributes = intColumnAttributesEnum
	End Property

	Public Property Get DefinedSize()
		DefinedSize = p_Column.DefinedSize
	End Property

	Public Property Let DefinedSize(lngDefinedSize)
		p_Column.DefinedSize = lngDefinedSize
	End Property

	Public Property Get Name()
		Name = p_Column.Name
	End Property

	Public Property Let Name(strName)
		p_Column.Name = strName
	End Property

	Public Property Get NumericScale()
		NumericScale = p_Column.NumericScale
	End Property

	Public Property Let NumericScale(bytNumericScale)
		p_Column.NumericScale = bytNumericScale
	End Property

	Public Property Get ParentCatalog()
		Set ParentCatalog = p_Column.ParentCatalog
	End Property

	Public Property Set ParentCatalog(objCatalog)
		Set p_Column.ParentCatalog = objCatalog
	End Property

	Public Property Get Precision()
		Precision = p_Column.Precision
	End Property

	Public Property Let Precision(lngPrecision)
		p_Column.Precision = lngPrecision
	End Property

	Public Property Get Properties()
		Set Properties = p_Column.Properties
	End Property

	Public Property Get RelatedColumn()
		RelatedColumn = p_Column.RelatedColumn
	End Property

	Public Property Let RelatedColumn(strRelatedColumn)
		p_Column.RelatedColumn = strRelatedColumn
	End Property

	Public Property Get SortOrder()
		SortOrder = p_Column.SortOrder
	End Property

	Public Property Let SortOrder(intSortOrderEnum)
		p_Column.SortOrder = intSortOrderEnum
	End Property

	Public Property Get ColumnType()
		ColumnType = p_Column.Type
	End Property

	Public Property Let ColumnType(intDataTypeEnum)
		p_Column.Type = intDataTypeEnum
	End Property

	Private Sub Class_Terminate()
		Set p_Column = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_Ext_Column.vbs" Then

End If
