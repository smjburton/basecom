Option Explicit

Class base_Database_DAO_Group
	Private p_Group

	Private Sub Class_Initialize()
		Set p_Group = CreateObject("DAO.Group.36")
	End Sub


	' Properties


	Public Property Get Name()
		Name = p_Group.Name
	End Property

	Public Property Let Name(strName)
		p_Group.Name = strName
	End Property

	Public Property Get PID()
		PID = p_Group.PID
	End Property

	Public Property Let PID(strPID)
		p_Group.PID = PID
	End Property

	Public Property Get Properties()
		Set Properties = p_Group.Properties
	End Property

	Public Default Property Get Users()
		Set Users = p_Group.Users
	End Property


	' Methods


	Public Function CreateUser() ' Optional params: [Name], [PID], [Password]
		Set CreateUser = p_Group.CreateUser()
	End Function

	Private Sub Class_Terminate()
		Set p_Group = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_DAO_Group.vbs" Then

End If
