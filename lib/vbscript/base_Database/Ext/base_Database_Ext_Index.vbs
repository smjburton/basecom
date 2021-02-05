Option Explicit

Class base_Database_Ext_Index
	Private p_Index
	
	Private Sub Class_Initialize()
		Set p_Index = CreateObject("ADOX.Index")
	End Sub


	' Properties


	Public Property Get Clustered()
		Clustered = p_Index.Clustered
	End Property

	Public Property Let Clustered(blnClustered)
		p_Index.Clustered = blnClustered
	End Property

	Public Property Get Columns()
		Set Columns = p_Index.Columns
	End Property

	Public Property Get IndexNulls()
		IndexNulls = p_Index.IndexNulls
	End Property

	Public Property Let IndexNulls(intAllowNullsEnum)
		p_Index.IndexNulls = intAllowNullsEnum
	End Property

	Public Property Get Name()
		Name = p_Index.Name
	End Property

	Public Property Let Name(strName)
		p_Index.Name = strName
	End Property

	Public Property Get PrimaryKey()
		PrimaryKey = p_Index.PrimaryKey
	End Property

	Public Property Let PrimaryKey(blnPrimaryKey)
		p_Index.PrimaryKey = blnPrimaryKey
	End Property

	Public Property Get Properties()
		Set Properties = p_Index.Properties
	End Property

	Public Property Get Unique()
		Unique = p_Index.Unique
	End Property

	Public Property Let Unique(blnUnique)
		p_Index.Unique = blnUnique
	End Property

	Private Sub Class_Terminate()
		Set p_Index = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_Ext_Index.vbs" Then

End If
