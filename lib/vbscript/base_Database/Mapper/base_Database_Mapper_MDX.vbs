Option Explicit

Class base_DB_MDX
	Private pConnection
	Private pCatalog
	Private pCellset

	Private pRowCount
	Private pColumnCount

	Private Sub Class_Initialize()
    		Set pConnection = CreateObject("ADODB.Connection.6.0")
    		Set pCatalog = CreateObject("ADOMD.Catalog.6.0")
    		Set pCellset = CreateObject("ADOMD.Cellset.6.0")

		pRowCount = 0
		pColumnCount = 0
	End Sub


	' Properties


	Public Property Get ColumnNumber(strColumn)
		Dim strHeader, _
			i

		For i = 0 To pColumnCount - 1
			strHeader = pCellset.Axes(0).Positions(i).Members(0).Caption

			If strHeader = strColumn Then
				ColumnNumber = i
				Exit For
			End If
		Next
	End Property

	Public Property Get ColumnHeader(intColumn)
		If intColumn <= (pColumnCount - 1) Then ColumnHeader = pCellset.Axes(0).Positions(intColumn).Members(0).Caption
	End Property

	Public Property Get ColumnCount()
		ColumnCount = pColumnCount
	End Property

	Public Property Get RowNumber(strRow)
		Dim strHeader, _
			i

		For i = 0 To pRowCount - 1
			strHeader = pCellset.Axes(1).Positions(i).Members(0).Caption

			If strHeader = strRow Then
				RowNumber = i
				Exit For
			End If
		Next
	End Property

	Public Property Get RowHeader(intRow)
		If intRow <= (pRowCount - 1) Then RowHeader = pCellset.Axes(1).Positions(intRow).Members(0).Caption
	End Property

	Public Property Get RowCount()
		RowCount = pRowCount
	End Property

	Public Default Property Get Data(varColumn, varRow)
		If TypeName(varColumn) = "String" Then
			varColumn = ColumnNumber(varColumn)
		End If

		If TypeName(varRow) = "String" Then
			varRow = RowNumber(varRow)
		End If

		Data = pCellset(varColumn, varRow)
	End Property

	Public Property Get Status()
		Status = pConnection.State
	End Property


	' Methods


	Public Sub Connect(strConnection)
		pConnection.Open strConnection
    		Set pCatalog.ActiveConnection = pConnection
    		Set pCellset.ActiveConnection = pCatalog.ActiveConnection
	End Sub

	Public Sub Query(strQuery)
		If Not pCellset.ActiveConnection Is Nothing Then
			If pCellset.State = 1 Then pCellset.Close()
    			pCellset.Source = strQuery
    			pCellset.Open()

			WScript.Echo pCellset.Axes(1).Positions(4).Members(1).Caption

			WScript.Echo pCellset(0, 3)

			pRowCount = pCellset.Axes(1).Positions.Count
			pColumnCount = pCellset.Axes(0).Positions.Count
		End If
	End Sub

	Public Sub Disconnect()

	End Sub

	Public Sub Flush()

	End Sub

	Public Sub ExportToCSV(strFilename)
		
	End Sub

	Public Sub ExportToExcel(strFilename)

	End Sub

	Private Sub Class_Terminate()
        	Set pConnection = Nothing
        	Set pCatalog = Nothing
        	Set pCellset = Nothing
	End Sub
End Class


If WScript.ScriptName = "base_DB_MDX.vbs" Then

End If
