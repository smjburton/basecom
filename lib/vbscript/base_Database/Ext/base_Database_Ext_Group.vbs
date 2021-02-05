Option Explicit

Class base_Database_Ext_Group
	Private p_Group

	Private Sub Class_Initialize()
		Set p_Group = CreateObject("ADOX.Group")
	End Sub


	' Properties


	Public Default Property Get Name()
		Name = p_Group.Name
	End Property

	Public Property Let Name(strName)
		p_Group.Name = strName
	End Property

	Public Property Get ParentCatalog()
		Set ParentCatalog = p_Group.ParentCatalog
	End Property

	Public Property Set ParentCatalog(objCatalog)
		Set p_Group.ParentCatalog = objCatalog
	End Property

	Public Property Get Properties()
		Set Properties = p_Group.Properties
	End Property

	Public Property Get Users()
		Set Users = p_Group.Users
	End Property


	' Methods


	Public Function GetPermissions(strName, intObjectTypeEnum) ' Optional params: [ObjectTypeId]) As RightsEnum
		GetPermissions = p_Group.GetPermissions(strName, intObjectTypeEnum)
	End Function

	Public Sub SetPermissions(strName, intObjectTypeEnum, intAction, intRightsEnum) ' Optional params: [Inherit As InheritTypeEnum = adInheritNone], [ObjectTypeId])
		p_Group.SetPermissions strName, intObjectTypeEnum, intAction, intRightsEnum
	End Sub

	Private Sub Class_Terminate()
		Set p_Group = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_Ext_Group.vbs" Then

End If
