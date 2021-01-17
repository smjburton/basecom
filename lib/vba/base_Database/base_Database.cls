Option Explicit


' Table Object (name, Column(...), Column(...), ...)

' result = conn.execute(ins)

' s = select([users])
' result = conn.execute(s)

' select([users, addresses]).where(users.c.id == addresses.c.user_id)



' insert()
' insert.values()
' values
' from(...)
' where(...).values(...)
' and(...)
' or(...)
' union(...)
' union_all(...)
' .asc() 
' .desc()
' count(...)
' average(...)
' sum(...)
' group_by(...)
' order_by(...)
' .limit(1).offset(1)
' having(...)
' .distinct()
' .like(...)
' .join() 
' .outerjoin() 
' .IsNull()
' .IsNotNull()
' .Top()
' .Min()
' .Max()
' update()
' update().values(...)
' delete()
' create_engine
' case(...)
' INSERT INTO SELECT / SELECT INTO
' ANY / ALL
' EXISTS
' BETWEEN

' Include "base_Database_MDX"
' Include "base_Datbase_SQL"

Sub Print(strText)
WScript.Echo strText
End Sub

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
