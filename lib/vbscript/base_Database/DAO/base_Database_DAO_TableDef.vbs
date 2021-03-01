Option Explicit

Class base_Database_DAO_TableDef
	Private p_TableDef

	Private Sub Class_Initialize()
		Set p_TableDef = CreateObject("DAO.TableDef.36")
	End Sub


	' Properties


	Public Property Get Attributes()
		Attributes = p_TableDef.Attributes
	End Property

	Public Property Let Attributes(lngAttributes)
		p_TableDef.Attributes = lngAttributes
	End Property

	Public Property Get ConflictTable()
		ConflictTable = p_TableDef.ConflictTable
	End Property

	Public Property Get Connect()
		Connect = p_TableDef.Connect
	End Property

	Public Property Let Connect(strConnect)
		p_TableDef.Connect = strConnect
	End Property

	Public Property Get DateCreated()
		If IsObject(p_TableDef.DateCreated) Then
			Set DateCreated = p_TableDef.DateCreated
		Else
			DateCreated = p_TableDef.DateCreated
		End If
	End Property

	Public Default Property Get Fields()
		Set Fields = p_TableDef.Fields
	End Property

	Public Property Get Indexes()
		Set Indexes = p_TableDef.Indexes
	End Property

	Public Property Get LastUpdated()
		If IsObject(p_TableDef.LastUpdated) Then
			Set LastUpdated = p_TableDef.LastUpdated
		Else
			LastUpdated = p_TableDef.LastUpdated
		End If
	End Property

	Public Property Get Name()
		Name = p_TableDef.Name
	End Property

	Public Property Let Name(strName)
		p_TableDef.Name = strName
	End Property

	Public Property Get Properties()
		Set Properties = p_TableDef.Properties
	End Property

	Public Property Get RecordCount()
		RecordCount = p_TableDef.RecordCount
	End Property

	Public Property Get ReplicaFilter()
		If IsObject(p_TableDef.ReplicaFilter) Then
			Set ReplicaFilter = p_TableDef.ReplicaFilter
		Else
			ReplicaFilter = p_TableDef.ReplicaFilter
		End If
	End Property

	Public Property Let ReplicaFilter(varReplicaFilter)
		p_TableDef.ReplicaFilter = varReplicaFilter
	End Property

	Public Property Set ReplicaFilter(varReplicaFilter)
		Set p_TableDef.ReplicaFilter = varReplicaFilter
	End Property

	Public Property Get SourceTableName()
		SourceTableName = p_TableDef.SourceTableName
	End Property

	Public Property Let SourceTableName(strSourceTableName)
		p_TableDef.SourceTableName = strSourceTableName
	End Property

	Public Property Get Updatable()
		Updatable = p_TableDef.Updatable
	End Property

	Public Property Get ValidationRule()
		ValidationRule = p_TableDef.ValidationRule
	End Property

	Public Property Let ValidationRule(strValidationRule)
		p_TableDef.ValidationRule = strValidationRule
	End Property

	Public Property Get ValidationText()
		ValidationText = p_TableDef.ValidationText
	End Property

	Public Property Let ValidationText(strValidationText)
		p_TableDef.ValidationText = strValidationText
	End Property


	' Methods


	Public Function CreateField() ' Optional params: [Name], [Type], [Size]
		Set CreateField = p_TableDef.CreateField()
	End Function

	Public Function CreateIndex() ' Optional params: [Name]
		Set CreateIndex = p_TableDef.CreateIndex()
	End Function

	Public Function CreateProperty() ' Optional params: [Name], [Type], [Value], [DDL]
		Set CreateProperty = p_TableDef.CreateProperty()
	End Function

	Public Function OpenRecordset() ' Optional params: [Type], [Options]
		Set OpenRecordset = p_TableDef.OpenRecordset()
	End Function

	Public Sub RefreshLink()
		p_TableDef.RefreshLink
	End Sub

	Private Sub Class_Terminate()
		Set p_TableDef = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_DAO_TableDef.vbs" Then

End If