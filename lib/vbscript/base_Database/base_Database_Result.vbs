Option Explicit

Class base_Database_Result
	Private p_Result 

	Private Sub Class_Initialize()
		Set p_Result = CreateObject("ADODB.Record")
	End Sub


	' Properties


	Public Property Get ActiveConnection()
		If IsObject(p_Result.ActiveConnection) Then
			Set ActiveConnection = p_Result.ActiveConnection 
		Else
			ActiveConnection = p_Result.ActiveConnection 
		End If
	End Property

	Public Property Let ActiveConnection(varActiveConnection)
		p_Result.ActiveConnection = varActiveConnection
	End Property

	Public Property Set ActiveConnection(varActiveConnection)
		Set p_Result.ActiveConnection = varActiveConnection
	End Property

	Public Default Property Get Fields()
		Set Fields = p_Result.Fields
	End Property

	Public Property Get Mode()
		p_Result.Mode
	End Property

	Public Property Let Mode(intConnectModeEnum)
		p_Result.Mode = intConnectModeEnum
	End Property

	Public Property Get ParentUrl()
		ParentUrl = p_Result.ParentURL 
	End Property

	Public Property Get Properties()
		Set Properties = p_Result.Properties 
	End Property

	Public Property Get RecordType()
		RecordType = p_Result.RecordType 
	End Property

	Public Property Get Source()
		If IsObject(p_Result.Source) Then
			Set Source = p_Result.Source
		Else
			Source = p_Result.Source
		End If
	End Property

	Public Property Let Source(varSource)
		p_Result.Source = varSource
	End Property

	Public Property Set Source(varSource)
		Set p_Result.Source = varSource
	End Property

	Public Property Get State()
		State = p_Result.State
	End Property


	' Methods


	Public Sub Cancel()
		p_Result.Cancel
	End Sub

	Public Sub Close()
		p_Result.Close
	End Sub

	Public Function CopyRecord() ' Optional params: [Source As String], [Destination As String], [UserName As String], [Password As String], [Options As CopyRecordOptionsEnum = adCopyUnspecified], [Async As Boolean = False]) As String
		CopyRecord = p_Result.CopyRecord()
	End Function

	Public Sub DeleteRecord() ' Optional params: Source As String], [Async As Boolean = False])
		p_Result.DeleteRecord
	End Sub

	Public Function GetChildren()
		Set GetChildren = p_Result.GetChildren()
	End Function

	Public Function MoveRecord() ' Optional params: Source As String], [Destination As String], [UserName As String], [Password As String], [Options As MoveRecordOptionsEnum = adMoveUnspecified], [Async As Boolean = False]) As String
		MoveRecord = p_Result.MoveRecord()
	End Function

	Public Sub Open() ' Optional params: Source], [ActiveConnection], [Mode As ConnectModeEnum = adModeUnknown], [CreateOptions As RecordCreateOptionsEnum = adFailIfNotExists], [Options As RecordOpenOptionsEnum = adOpenRecordUnspecified], [UserName As String], [Password As String])
		p_Result.Open
	End Sub

	Private Sub Class_Terminate()
		Set p_Result = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_Result.vbs" Then

End If
