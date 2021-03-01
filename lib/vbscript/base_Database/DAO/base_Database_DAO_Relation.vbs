Option Explicit

Class base_Database_DAO_Relation
	Private p_Relation

	Private Sub Class_Initialize()
		Set p_Relation = CreateObject("DAO.Relation.36")
	End Sub


	' Properties


	Public Property Get Attributes()
		Attributes = p_Relation.Attributes
	End Property

	Public Property Let Attributes(lngAttributes)
		p_Relation.Attributes = lngAttributes
	End Property

	Public Default Property Get Fields()
		Set Fields = p_Relation.Fields
	End Property

	Public Property Get ForeignTable()
		ForeignTable = p_Relation.ForeignTable
	End Property

	Public Property Let ForeignTable(strForeignTable)
		p_Relation.ForeignTable = strForeignTable
	End Property

	Public Property Get Name()
		Name = p_Relation.Name
	End Property

	Public Property Let Name(strName)
		p_Relation.Name = strName
	End Property

	Public Property Get PartialReplica()
		PartialReplica = p_Relation.PartialReplica
	End Property

	Public Property Let PartialReplica(blnPartialReplica)
		p_Relation.PartialReplica = blnPartialReplica
	End Property

	Public Property Get Properties()
		Set Properties = p_Relation.Properties
	End Property

	Public Property Get Table()
		Table = p_Relation.Table
	End Property

	Public Property Let Table(strTable)
		p_Relation.Table = strTable
	End Property


	' Methods


	Public Function CreateField() ' Optional params: [Name], [Type], [Size]
		Set CreateField = p_Relation.CreateField()
	End Function

	Private Sub Class_Terminate()
		Set p_Relation = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_DAO_Relation.vbs" Then

End If
