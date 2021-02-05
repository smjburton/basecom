Option Explicit

Class base_Database_Ext_Catalog
	Private p_Catalog

	Private Sub Class_Initialize()
		Set p_Catalog = CreateObject("ADOX.Catalog")
	End Sub


	' Properties


	Public Property Get ActiveConnection()
		If IsObjecT(p_Catalog.ActiveConnection) Then
			Set ActiveConnection = p_Catalog.ActiveConnection
		Else
			ActiveConnection = p_Catalog.ActiveConnection
		End If
	End Property

	Public Property Let ActiveConnection(varConnection)
		p_Catalog.ActiveConnection = varConnection
	End Property

	Public Property Set ActiveConnection(varConnection)
		Set p_Catalog.ActiveConnection = varConnection
	End Property

	Public Property Get Groups()
		Set Groups = p_Catalog.Groups
	End Property

	Public Property Get Procedures()
		Set Procedures = p_Catalog.Procedures
	End Property

	Public Default Property Get Tables()
		Set Tables = p_Catalog.Tables
	End Property

	Public Property Get Users()
		Set Users = p_Catalog.Users
	End Property

	Public Property Get Views()
		Set Views = p_Catalog.Views
	End Property


	' Methods


	Public Function Create(strConnectString)
		Set Create = p_Catalog.Create(strConnectString)
	End Function

	Public Function GetObjectOwner(strObjectName, intObjectTypeEnum) ' Optional params: [ObjectTypeId]) As String
		GetObjectOwner = p_Catalog.GetObjectOwner(strObjectName, intObjectTypeEnum)
	End Function

	Public Sub SetObjectOwner(strObjectName, intObjectTypeEnum, strUserName) ' Optional params: [ObjectTypeId])
		p_Catalog.SetObjectOwner strObjectName, intObjectTypeEnum, strUserName
	End Sub

	Private Sub Class_Terminate()
		Set p_Catalog = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_Ext_Catalog.vbs" Then

End If
