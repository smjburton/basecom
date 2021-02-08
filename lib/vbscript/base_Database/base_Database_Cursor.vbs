Option Explicit

Class base_Database_Cursor
	Private p_Cursor

	Private Sub Class_Initialize()
		Set p_Cursor = CreateObject("ADODB.Command")
	End Sub


	' Properties


	Public Property Get ActiveConnection()
		Set ActiveConnection = p_Cursor.ActiveConnection
	End Property

	Public Property Set ActiveConnection(objConnection)
		Set p_Cursor.ActiveConnection = objConnection
	End Property

	Public Property Get CursorType()
		CursorType = p_Cursor.CommandType
	End Property

	Public Property Let CursorType(intCursorTypeEnum)
		p_Cursor.CommandType = intCursorTypeEnum
	End Property

	Public Property Get Dialect()
		Dialect = p_Cursor.Dialect
	End Property

	Public Property Let Dialect(strDialect)
		p_Cursor.Dialect = strDialect
	End Property

	Public Property Get Name()
		Name = p_Cursor.Name
	End Property

	Public Property Let Name(strName)
		p_Cursor.Name = strName
	End Property

	Public Property Get NamedParameters()
		NamedParameters = p_Cursor.NamedParameters
	End Property

	Public Property Let NamedParameters(blnNamedParameters)
		p_Cursor.NamedParameters = blnNamedParameters
	End Property

	Public Property Get Parameters()
		Parameters = p_Cursor.Parameters
	End Property

	Public Property Get Prepared()
		Prepared = p_Cursor.Prepared
	End Property

	Public Property Let Prepared(blnPrepared)
		p_Cursor.Prepared = blnPrepared
	End Property

	Public Property Get Properties()
		Set Properties = p_Cursor.Properties
	End Property

	Public Property Get State()
		State = p_Cursor.State
	End Property

	Public Property Get Stream()
		If IsObject(p_Cursor.CommandStream) Then
			Set Stream = p_Cursor.CommandStream
		Else
			Stream = p_Cursor.CommandStream
		End If
	End Property

	Public Property Let Stream(varStream)
		p_Cursor.CommandStream = varStream
	End Property

	Public Property Set Stream(varStream)
		Set p_Cursor.CommandStream = varStream
	End Property

	Public Property Get Text()
		Text = p_Cursor.CommandText
	End Property

	Public Property Let Text(strText)
		p_Cursor.CommandText = strText
	End Property

	Public Property Get Timeout()
		CommandTimeout = p_Cursor.CommandTimeout
	End Property

	Public Property Let Timeout(lngTimeout)
		p_Cursor.CommandTimeout = lngCommandTimeout
	End Property


	' Methods


	Public Sub Cancel()
		p_Cursor.Cancel
	End Sub

	Public Function CreateParameter() ' Optional params: [Name As String], [Type As DataTypeEnum = adEmpty], [Direction As ParameterDirectionEnum = adParamInput], [Size As Long], [Value]) As Parameter
		Set CreateParameter = p_Cursor.CreateParameter()
	End Function

	Public Function Execute(strCursorText) ' Optional params: [RecordsAffected], [Parameters], [Options As Long = -1]) As Recordset
		p_Cursor.CommandText = strCursorText
		Set Execute = p_Cursor.Execute()
	End Function

	Private Sub Class_Terminate()
		Set p_Cursor = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_Cursor.vbs" Then

End If