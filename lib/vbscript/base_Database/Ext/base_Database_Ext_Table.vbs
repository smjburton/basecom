Option Explicit

Class base_Database_Ext_Table
	Private p_Table

	Private Sub Class_Initialize()
		Set p_Table = CreateObject("ADOX.Table")
	End Sub


	' Properties


	Public Default Property Get Columns()
		Set Columns = p_Table.Columns
	End Property

	Public Property Get DateCreated()
		If IsObject(p_Table.DateCreated) Then
			Set DateCreated = p_Table.DateCreated
		Else
			DateCreated = p_Table.DateCreated
		End If
	End Property

	Public Property Get DateModified()
		If IsObject(p_Table.DateCreated) Then
			Set DateModified = p_Table.DateModified
		Else
			DateModified = p_Table.DateModified
		End If
	End Property

	Public Property Get Indexes()
		Set Indexes = p_Table.Indexes
	End Property

	Public Property Get Keys()
		Set Keys = p_Table.Keys
	End Property

	Public Property Get Name()
		Name = p_Table.Name
	End Property

	Public Property Let Name(strName)
		p_Table.Name = strName
	End Property

	Public Property Get ParentCatalog()
		Set ParentCatalog = p_Table.ParentCatalog
	End Property

	Public Property Set ParentCatalog(objCatalog)
		Set p_Table.ParentCatalog = objCatalog
	End Property

	Public Property Get Properties()
		Set Properties = p_Table.Properties
	End Property

	Public Property Get TableType()
		TableType = p_Table.Type
	End Property

	Private Sub Class_Terminate()
		Set p_Table = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_Ext_Table.vbs" Then

End If
