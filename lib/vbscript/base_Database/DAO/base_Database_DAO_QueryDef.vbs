Option Explicit

Class base_Database_DAO_QueryDef
	Private p_QueryDef

	Private Sub Class_Initialize()
		Set p_QueryDef = CreateObject("DAO.QueryDef.36")
	End Sub


	' Properties


	Public Property Get CacheSize()
		CacheSize = p_QueryDef.CacheSize
	End Property

	Public Property Let CacheSize(lngCacheSize)
		p_QueryDef.CacheSize = lngCacheSize
	End Property

	Public Property Get Connect()
		Connect = p_QueryDef.Connect
	End Property

	Public Property Let Connect(strConnect)
		p_QueryDef.Connect = strConnect
	End Property

	Public Property Get DateCreated()
		If IsObject(p_QueryDef.DateCreated) Then
			Set DateCreated = p_QueryDef.DateCreated
		Else
			DateCreated = p_QueryDef.DateCreated
		End If
	End Property

	Public Property Get Fields()
		Set Fields = p_QueryDef.Fields
	End Property

	Public Property Get LastUpdated()
		If IsObject(p_QueryDef.LastUpdated) Then
			Set LastUpdated = p_QueryDef.LastUpdated
		Else
			LastUpdated = p_QueryDef.LastUpdated
		End If
	End Property

	Public Property Get MaxRecords()
		MaxRecords = p_QueryDef.MaxRecords
	End Property

	Public Property Let MaxRecords(lngMaxRecords)
		p_QueryDef.MaxRecords = lngMaxRecords
	End Property

	Public Property Get Name()
		Name = p_QueryDef.Name
	End Property

	Public Property Let Name(strName)
		p_QueryDef.Name = strName
	End Property

	Public Property Get ODBCTimeout()
		ODBCTimeout = p_QueryDef.ODBCTimeout
	End Property

	Public Property Let ODBCTimeout(intODBCTimeout)
		p_QueryDef.ODBCTimeout = intODBCTimeout
	End Property

	Public Default Property Get Parameters()
		Set Parameters = p_QueryDef.Parameters
	End Property

	Public Property Get Prepare()
		If IsObject(p_QueryDef.Prepare) Then
			Set Prepare = p_QueryDef.Prepare
		Else
			Prepare = p_QueryDef.Prepare
		End If
	End Property

	Public Property Let Prepare(varPrepare)
		p_QueryDef.Prepare = varPrepare
	End Property

	Public Property Set Prepare(varPrepare)
		Set p_QueryDef.Prepare = varPrepare
	End Property

	Public Property Get Properties()
		Set Properties = p_QueryDef.Properties
	End Property

	Public Property Get RecordsAffected()
		RecordsAffected = p_QueryDef.RecordsAffected
	End Property

	Public Property Get ReturnsRecords()
		ReturnsRecords = p_QueryDef.ReturnsRecords
	End Property

	Public Property Let ReturnsRecords(blnReturnsRecords)
		p_QueryDef.ReturnsRecords = blnReturnsRecords
	End Property

	Public Property Get SQL()
		SQL = p_QueryDef.SQL
	End Property

	Public Property Let SQL(strSQL)
		p_QueryDef.SQL = strSQL
	End Property

	Public Property Get StillExecuting()
		StillExecuting = p_QueryDef.StillExecuting
	End Property

	Public Property Get QueryType()
		QueryType = p_QueryDef.Type
	End Property

	Public Property Get Updatable()
		Updatable = p_QueryDef.Updatable
	End Property


	' Methods


	Public Sub Cancel()
		p_QueryDef.Cancel
	End Sub

	Public Sub Close()
		p_QueryDef.Close
	End Sub

	Public Function CreateProperty() ' Optional params: [Name], [Type], [Value], [DDL]
		Set CreateProperty = p_QueryDef.CreateProperty()
	End Function

	Public Sub Execute() ' Optional params: [Options]
		p_QueryDef.Execute
	End Sub

	Public Function OpenRecordset() ' Optional params: [Type], [Options], [LockEdit]
		Set OpenRecordset = p_QueryDef.OpenRecordset()
	End Function

	Private Sub Class_Terminate()
		Set p_QueryDef = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_DAO_QueryDef.vbs" Then

End If
