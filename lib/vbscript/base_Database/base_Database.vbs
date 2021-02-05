Option Explicit

Class base_Database
	Private pConnection, _
			pRecordset


	' Constructor


	Private Sub Class_Initialize()
		Set pConnection = CreateObject("ADODB.Connection.6.0")
		Set pRecordset = CreateObject("ADODB.Recordset.6.0")
	End Sub


	' Properties


	Public Property Get Connection()
		Set Connection = pConnection
	End Property

	Public Property Get Recordset()
		Set Recordset = pRecordset
	End Property

	Public Property Get Fields(strField)
		Fields = pRecordset.Fields(strField)
	End Property


	' Methods


	Public Sub CreateEngine(strConnection)
		pConnection.Open strConnection
		pRecordSet.ActiveConnection = pConnection
	End Sub

	Public Sub Execute(strSql)
		pRecordset.Open strSql, pConnection
	End Sub

	Public Sub SqlSelect(strSelect, strTable)
		Me.Execute("SELECT " & strSelect & " FROM " & strTable & ";")
	End Sub


	' Helper Methods


	Private Function baseDatabase(varDatabase)

	End Function


	' Deconstructor


	Private Sub Class_Terminate()
		' pConnection.Close()
		Set pConnection = Nothing
		' pRecordset.Close()
		Set pRecordset = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database.vbs" Then
	Dim objDb, _
		strSql, _
		strConnection

	Set objDb = New base_Database

	strConnection = "DSN=PostgreSQL35W;Uid=postgres;Pwd=postgres"

	With objDb
		.CreateEngine strConnection
		.SqlSelect "*", "users" 
		Print .Fields(1)
	End With
End If
