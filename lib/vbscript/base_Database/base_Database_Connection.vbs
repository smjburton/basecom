Option Explicit

' Include "base_Database_Cursor"

Class base_Database_Connection
	Private p_Connection

	Private Sub Class_Initialize()
		Set p_Connection = CreateObject("ADODB.Connection")
	End Sub


	' Properties


	Public Property Get Attributes()
		Attributes = p_Connection.Attributes
	End Property

	Public Property Let Attributes(lngAttributes)
		p_Connection.Attributes = lngAttributes
	End Property

	Public Property Get CommandTimeout()
		CommandTimeout = p_Connection.CommandTimeout
	End Property

	Public Property Let CommandTimeout(lngCommandTimeout)
		p_Connection.CommandTimeout = lngCommandTimeout
	End Property

	Public Property Get ConnectionString()
		ConnectionString = p_Connection.ConnectionString
	End Property

	Public Property Let ConnectionString(strConnectionString)
		p_Connection.ConnectionString = strConnectionString
	End Property

	Public Property Get ConnectionTimeout()
		ConnectionTimeout = p_Connection.ConnectionTimeout
	End Property

	Public Property Let ConnectionTimeout(lngConnectionTimeout)
		p_Connection.ConnectionTimeout = lngConnectionTimeout
	End Property

	Public Property Get CursorLocation()
		CursorLocation = p_Connection.CursorLocation
	End Property

	Public Property Let CursorLocation(intCursorLocationEnum)
		p_Connection.CursorLocation = intCursorLocationEnum
	End Property

	Public Property Get DefaultDatabase()
		DefaultDatabase = p_Connection.DefaultDatabase
	End Property

	Public Property Let DefaultDatabase(strDefaultDatabase)
		p_Connection.DefaultDatabase = strDefaultDatabase
	End Property

	Public Property Get Errors()
		Set Errors = p_Connection.Errors
	End Property

	Public Property Get IsolationLevel()
		IsolationLevel = p_Connection.IsolationLevel
	End Property

	Public Property Let IsolationLevel(intIsolationLevelEnum)
		p_Connection.IsolationLevel = intIsolationLevelEnum
	End Property

	Public Property Get Mode()
		Mode = p_Connection.Mode
	End Property

	Public Property Let Mode(intConnectModeEnum)
		p_Connection.Mode = intConnectModeEnum
	End Property

	Public Property Get Properties(objProperties)
		Set p_Connection.Properties = objProperties
	End Property

	Public Property Get Provider()
		Provider = p_Connection.Provider
	End Property

	Public Property Let Provider(strProvider)
		p_Connection.Provider = strProvider
	End Property

	Public Property Get State()
		State = p_Connection.State
	End Property

	Public Property Get Version()
		Version = p_Connection.Version
	End Property


	' Methods


	Public Function BeginTrans()
		BeginTrans  = p_Connection.BeginTrans()
	End Function

	Public Sub Cancel()
		p_Connection.Cancel
	End Sub

	Public Sub Close()
		p_Connection.Close
	End Sub

	Public Sub CommitTrans()
		p_Connection.CommitTrans
	End Sub

	Public Function Execute(strCommandText) ' Optional params: [RecordsAffected], [Options As Long = -1]) As Recordset
		Set Execute = p_Connection.Execute(strCommandText)
	End Function

	Public Sub Open() ' Optional params: [ConnectionString As String], [UserID As String], [Password As String], [Options As Long = -1])
		p_Connection.Open
	End Sub

	Public Function OpenSchema(intSchema) ' Optional params: [Restrictions], [SchemaID]) As Recordset
		Set OpenSchema = p_Connection.OpenSchema(intSchema)
	End Function

	Public Sub RollbackTrans()
		p_Connection.RollbackTrans
	End Sub


	' Events


	' BeginTransComplete
	' CommitTransComplete
	' ConnectComplete
	' Disconnect
	' ExecuteComplete
	' InfoMessage
	' RollbackTransComplete
	' WillConnect
	' WillExecute

	Private Sub Class_Terminate()
		Set p_Connection = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_Connection.vbs" Then

End If
