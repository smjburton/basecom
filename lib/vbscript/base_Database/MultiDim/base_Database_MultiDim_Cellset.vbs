Option Explicit

Class base_Database_MultiDim_Cellset
	Private p_Cellset

	Private Sub Class_Initialize()
		Set p_Cellset = CreateObject("ADOMD.Cellset")
	End Sub


	' Properties


	Public Property Get ActiveConnection()
		Set ActiveConnection = p_Cellset.ActiveConnection
	End Property

	Public Property Set ActiveConnection(objActiveConnection)
		Set p_Cellset.ActiveConnection = objActiveConnection
	End Property

	Public Property Get Axes()
		Set Axes = p_Cellset.Axes
	End Property

	Public Property Get FilterAxis()
		Set FilterAxis = p_Cellset.FilterAxis
	End Property

	Public Property Get Item(varIndex())
		Set Item = p_Cellset.Item(varIndex())
	End Property

	Public Property Get Properties()
		Set Properties = p_Cellset.Properties
	End Property

	Public Property Get Source()
		If IsObject(p_Cellset.Source) Then
			Set Source = p_Cellset.Source
		Else
			Source = p_Cellset.Source
		End If
	End Property

	Public Property Let Source(varSource)
		p_Cellset.Source = varSources
	End Property

	Public Property Set Source(varSource)
		Set p_Cellset.Source = varSource
	End Property

	Public Property Get State()
		State = p_Cellset.State
	End Property


	' Methods


	Public Sub Close()
		p_Cellset.Close
	End Sub

	Public Sub Open() ' Optional params: [DataSource], [ActiveConnection])
		p_Cellset.Open
	End Sub

	Private Sub Class_Terminate()
		Set p_Cellset = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_MultiDim_Cellset.vbs" Then

End If
