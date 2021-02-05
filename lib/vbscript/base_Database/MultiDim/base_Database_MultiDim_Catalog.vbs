Option Explicit

Class base_Database_MultiDim_Catalog
	Private p_Catalog

	Private Sub Class_Initialize()
		Set p_Catalog = CreateObject("ADOMD.Catalog")
	End Sub


	' Properties


	Public Property Get ActiveConnection()
		Set ActiveConnection = p_Catalog.ActiveConnection
	End Property

	Public Property Set ActiveConnection(objActiveConnection)
		Set p_Catalog.ActiveConnection = objActiveConnection
	End Property

	Public Property Get CubeDefs()
		Set CubeDefs = p_Catalog.CubeDefs
	End Property

	Public Property Get Name()
		Name = p_Catalog.Name
	End Property

	Private Sub Class_Terminate()
		Set p_Catalog = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_MultiDim_Catalog.vbs" Then

End If
