Option Explicit

Class base_Database_Ext_User
	Private p_User

	Private Sub Class_Initialize()
		Set p_User = CreateObject("ADOX.User")
	End Sub


	' Properties


	Public Property Get Groups()
		Set Groups = p_User.Groups
	End Property

	Public Default Property Get Name()
		Name = p_User.Name
	End Property

	Public Property Let Name(strName)
		p_User.Name = strName
	End Property

	Public Property Get ParentCatalog()
		Set ParentCatalog = p_User.ParentCatalog
	End Property

	Public Property Set ParentCatalog(objCatalog)
		Set p_User.ParentCatalog = objCatalog
	End Property

	Public Property Get Properties()
		Set Properties = p_User.Properties
	End Property


	' Methods


	Public Sub ChangePassword(strOldPassword, strNewPassword)
		p_User.ChangePassword strOldPassword, strNewPassword
	End Sub

	Public Function GetPermissions(strName, intObjectTypeEnum) ' Optional params: [ObjectTypeId]) As RightsEnum
		GetPermissions = p_User.GetPermissions(strName, intObjectTypeEnum)
	End Function

	Public Sub SetPermissions(strName, intObjectTypeEnum, intActionEnum, intRightsEnum) ' Optional params: [Inherit As InheritTypeEnum = adInheritNone], [ObjectTypeId])
		p_User.SetPermissions strName, intObjectTypeEnum, intActionEnum, intRightsEnum
	End Sub

	Private Sub Class_Terminate()
		Set p_User = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_Ext_User.vbs" Then

End If
