Option Explicit

Class base_Database_Ext_Key
	Private p_Key

	Private Sub Class_Initialize()
		Set p_Key = CreateObject("ADOX.Key")
	End Sub


	' Properties


	Public Property Get Columns()
		Set Columns = p_Key.Columns
	End Property

	Public Property Get DeleteRule()
		DeleteRule = p_Key.DeleteRule
	End Property

	Public Property Let DeleteRule(intRuleEnum)
		p_Key.DeleteRule = intRuleEnum
	End Property

	Public Property Get Name()
		Name = p_Key.Name
	End Property

	Public Property Let Name(strName)
		p_Key.Name = strName
	End Property

	Public Property Get RelatedTable()
		RelatedTable = p_Key.RelatedTable
	End Property

	Public Property Let RelatedTable(strRelatedTable)
		p_Key.RelatedTable = strRelatedTable
	End Property

	Public Property Get KeyType()
		KeyType = p_Key.Type
	End Property

	Public Property Let KeyType(intKeyTypeEnum)
		p_Key.Type = intKeyTypeEnum
	End Property

	Public Property Get UpdateRule()
		UpdateRule = p_Key.UpdateRule
	End Property

	Public Property Let UpdateRule(intRuleEnum)
		p_Key.UpdateRule = intRuleEnum
	End Property

	Private Sub Class_Terminate()
		Set p_Key = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_Ext_Key.vbs" Then

End If
