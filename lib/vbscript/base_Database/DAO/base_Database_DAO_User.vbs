Option Explicit

Class base_Database_DAO_User
	Private p_User

	Private Sub Class_Initialize()
		Set p_User = CreateObject("DAO.User.36")
	End Sub


	' Properties


	Public Default Property Get Groups()
		Set Groups = p_User.Groups
	End Property

	Public Property Get Name()
		Name = p_User.Name
	End Property

	Public Property Let Name(strName)
		p_User.Name = strName
	End Property

	Public Property Get Password()
		Password = p_User.Password
	End Property

	Public Property Let Password(strPassword)
		p_User.Password = strPassword
	End Property

	Public Property Get PID()
		PID = p_User.PID
	End Property

	Public Property Let PID(strPID)
		p_User.PID = strPID
	End Property

	Public Property Get Properties()
		Set Properties = p_User.Properties
	End Property


	' Methods


	Public Function CreateGroup() ' Optional params: [Name], [PID] As Group
		CreateGroup = p_User.CreateGroup()
	End Function

	Public Sub NewPassword(strOldPassword, strNewPassword)
		p_User.NewPassword strOldPassword, strNewPassword
	End Sub

	Private Sub Class_Terminate()
		Set p_User = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_DAO_User.vbs" Then

End If
