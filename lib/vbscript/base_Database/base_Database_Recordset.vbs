Option Explicit

' https://bytes.com/topic/access/answers/890865-how-properly-open-recordset-ado
' http://www.w3schools.com/asp/met_rs_open.asp

Class base_DB_Recordset
	Private pRecordset

	Private Sub Class_Initialize()
		Set pRecordset = CreateObject("ADODB.Recordset")
	End Sub


	' Properties


	Public Property Get AbsolutePage()
		Set AbsolutePage = pRecordset.AbsolutePage
	End Property

	Public Property Get AbsolutePosition()
		Set AbsolutePosition = pRecordset.AbsolutePosition
	End Property

	Public Property Get ActiveCommand()
		Set ActiveCommand = pRecordset.ActiveCommand
	End Property

	Public Property Get ActiveConnection()
		If IsObject(pRecordset.ActiveConnection) Then
			Set ActiveConnection = pRecordset.ActiveConnection
		Else
			ActiveConnection = pRecordset.ActiveConnection
		End If
	End Property

	Public Property Get BOF()
		BOF = pRecordset.BOF
	End Property

	Public Property Get Bookmark()
		If IsObject(pRecordset.Bookmark) Then
			Set Bookmark = pRecordset.Bookmark
		Else
			Bookmark = pRecordset.Bookmark
		End If 
	End Property

	Public Property Get CacheSize()
		CacheSize = pRecordset.CacheSize
	End Property

	Public Property Get CursorLocation()
		Set CursorLocation = pRecordset.CursorLocation
	End Property

	Public Property Get CursorType()
		Set CursorType = pRecordset.CursorType
	End Property

	Public Property Get DataMember()
		DataMember = pRecordset.DataMember
	End Property

	Public Property Get DataSource()
		If IsObject(pRecordset.DataSource) Then
			Set DataSource = pRecordset.DataSource
		Else
			DataSource = pRecordset.DataSource
		End If 
	End Property

	Public Property Get EditMode()
		Set EditMode = pRecordset.EditMode
	End Property

	Public Property Get EOF()
		EOF = pRecordset.EOF
	End Property

	Public Property Get Fields()
		Set Fields = pRecordset.Fields
	End Property

	Public Property Get Filter()
		If IsObject(pRecordset.Filter) Then
			Set Filter = pRecordset.Filter
		Else
			Filter = pRecordset.Filter
		End If 
	End Property

	Public Property Get Index()
		Index = pRecordset.Index
	End Property

	Public Property Get LockType()
		Set LockType = pRecordset.LockType
	End Property

	Public Property Get MarshalOptions()
		Set MarshalOptions = pRecordset.MarshalOptions
	End Property

	Public Property Get MaxRecords()
		MaxRecords = pRecordset.MaxRecords
	End Property

	Public Property Get PageCount()
		PageCount = pRecordset.PageCount
	End Property

	Public Property Get PageSize()
		PageSize = pRecordset.PageSize
	End Property

	Public Property Get Properties()
		Set Properties = pRecordset.Properties
	End Property

	Public Property Get RecordCount()
		RecordCount = pRecordset.RecordCount
	End Property

	Public Property Get Sort()
		Sort = pRecordset.Sort
	End Property

	Public Property Get Source()
		If IsObject(pRecordset.Source) Then
			Set Source = pRecordset.Source
		Else
			Source = pRecordset.Source
		End If 
	End Property

	Public Property Get State()
		State = pRecordset.State
	End Property

	Public Property Get Status()
		Status = pRecordset.Status
	End Property

	Public Property Get StayInSync()
		StayInSync = pRecordset.StayInSync
	End Property


	' Methods


	' Public Sub AddNew([FieldList], [Values])

	' End Sub

	Public Sub Cancel()
		pRecordset.Cancel()
	End Sub

	' Public Sub CancelBatch([AffectRecords As AffectEnum = adAffectAll])

	' End Sub

	Public Sub CancelUpdate()
		pRecordset.CancelUpdate()
	End Sub

	' Public Function Clone([LockType As LockTypeEnum = adLockUnspecified]) 

	' End Function

	Public Sub Close()
		pRecordset.Close()
	End Sub

	' Public Function CompareBookmarks(Bookmark1, Bookmark2)

	' End Function

	' Public Sub Delete([AffectRecords As AffectEnum = adAffectCurrent])

	' End Sub

	' Public Sub Find(Criteria As String, [SkipRecords As Long], [SearchDirection As SearchDirectionEnum = adSearchForward], [Start])

	' End Sub

	' Public Function GetRows([Rows As Long = -1], [Start], [Fields])

	' End Function

	' Public Function GetString([StringFormat As StringFormatEnum = adClipString], [NumRows As Long = -1], [ColumnDelimeter As String], [RowDelimeter As String], [NullExpr As String])

	' End Function
 
	' Public Sub Move(NumRecords As Long, [Start])

	' End Sub

	Public Sub MoveFirst()
		pRecordset.MoveFirst()
	End Sub

	Public Sub MoveLast()
		pRecordset.MoveLast()
	End Sub

	Public Sub MoveNext()
		pRecordset.MoveNext()
	End Sub

	Public Sub MovePrevious()
		pRecordset.MovePrevious()
	End Sub

	' Public Function NextRecordset([RecordsAffected])

	' End Function

	' Public Sub Open([Source], [ActiveConnection], [CursorType As CursorTypeEnum = adOpenUnspecified], [LockType As LockTypeEnum = adLockUnspecified], [Options As Long = -1])

	' End Sub

	' Public Sub Requery([Options As Long = -1])

	' End Sub

	' Public Sub Resync([AffectRecords As AffectEnum = adAffectAll], [ResyncValues As ResyncEnum = adResyncAllValues])

	' End Sub

	' Public Sub Save([Destination], [PersistFormat As PersistFormatEnum = adPersistADTG])

	' End Sub

	' Public Sub Seek(KeyValues, [SeekOption As SeekEnum = adSeekFirstEQ])

	' End Sub

	' Public Function Supports(CursorOptions As CursorOptionEnum)

	' End Function
 
	' Public Sub Update([Fields], [Values])

	' End Sub

	' Public Sub UpdateBatch([AffectRecords As AffectEnum = adAffectAll])

	' End Sub


	Private Sub Class_Terminate()
		Set pRecordset = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_DB_Recordset.vbs" Then
	Dim recordset
	Set recordset = New base_DB_Recordset

	WScript.Echo TypeName(recordset.AbsolutePage)
	WScript.Echo TypeName(recordset.AbsolutePosition)
	WScript.Echo TypeName(recordset.ActiveCommand)
	WScript.Echo TypeName(recordset.ActiveConnection)
	WScript.Echo TypeName(recordset.BOF)
	WScript.Echo TypeName(recordset.Bookmark)
	WScript.Echo TypeName(recordset.CacheSize)
	WScript.Echo TypeName(recordset.CursorLocation)
	WScript.Echo TypeName(recordset.CursorType)
	WScript.Echo TypeName(recordset.DataMember)
	WScript.Echo TypeName(recordset.DataSource)
	WScript.Echo TypeName(recordset.EditMode)
	WScript.Echo TypeName(recordset.EOF)
	WScript.Echo TypeName(recordset.Fields)
	WScript.Echo TypeName(recordset.Filter)
	WScript.Echo TypeName(recordset.Index)
	WScript.Echo TypeName(recordset.LockType)
	WScript.Echo TypeName(recordset.MarshalOptions)
	WScript.Echo TypeName(recordset.MaxRecords)
	WScript.Echo TypeName(recordset.PageCount)
	WScript.Echo TypeName(recordset.PageSize)
	WScript.Echo TypeName(recordset.Properties)
	WScript.Echo TypeName(recordset.RecordCount)
	WScript.Echo TypeName(recordset.Sort)
	WScript.Echo TypeName(recordset.Source)
	WScript.Echo TypeName(recordset.State)
	WScript.Echo TypeName(recordset.Status)
	WScript.Echo TypeName(recordset.StayInSync)
End If
