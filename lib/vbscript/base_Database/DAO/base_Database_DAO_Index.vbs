Option Explicit

Class base_Database_DAO_Index
	Private p_Index

	Private Sub Class_Initialize()
		Set p_Index = CreateObject("DAO.Index.36")
	End Sub


	' Properties


	Public Property Get Clustered()
		Clustered = p_Index.Clustered
	End Property

	Public Property Let Clustered(blnClustered)
		p_Index.Clustered = blnClustered
	End Property

	Public Property Get DistinctCount()
		DistinctCount = p_Index.DistinctCount
	End Property

	Public Property Get Fields()
		Fields = p_Index.Fields
	End Property

	Public Property Let Fields(varFields)
		p_Index.Fields = varFields
	End Property

	Public Property Set Fields(varFields)
		Set p_Index.Fields = varFields
	End Property

	Public Property Get Foreign()
		Foreign = p_Index.Foreign
	End Property

	Public Property Get IgnoreNulls()
		IgnoreNulls = p_Index.IgnoreNulls
	End Property

	Public Property Let IgnoreNulls(blnIgnoreNulls)
		p_Index.IgnoreNulls = blnIgnoreNulls
	End Property

	Public Property Get Name()
		Name = p_Index.Name
	End Property

	Public Property Let Name(strName)
		p_Index.Name = strName
	End Property

	Public Property Get Primary()
		Primary = p_Index.Primary
	End Property

	Public Property Let Primary(blnPrimary)
		p_Index.Primary = blnPrimary
	End Property

	Public Property Get Properties()
		Set Properties = p_Index.Properties
	End Property

	Public Property Get Required()
		Required = p_Index.Required
	End Property

	Public Property Let Required(blnRequired)
		p_Index.Required = blnRequired
	End Property

	Public Property Get Unique()
		Unique = p_Index.Unique
	End Property

	Public Property Let Unique(blnUnique)
		p_Index.Unique = blnUnique
	End Property


	' Methods


	Public Function CreateField() ' Optional params: [Name], [Type], [Size]
		Set CreateField = p_Index.CreateField()
	End Function

	Public Function CreateProperty() ' Optional params: [Name], [Type], [Value], [DDL]
		Set CreateProperty = p_Index.CreateProperty()
	End Function

	Private Sub Class_Terminate()
		Set p_Index = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_DAO_Index.vbs" Then

End If
