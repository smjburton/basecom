Option Explicit

Class base_DB_Recordset
	Private p_Resultset

	Private Sub Class_Initialize()
		Set p_Resultset = CreateObject("ADODB.Recordset")
	End Sub


	' Properties


	Public Property Get AbsolutePage()
		Set AbsolutePage = p_Resultset.AbsolutePage
	End Property

	Public Property Get AbsolutePosition()
		Set AbsolutePosition = p_Resultset.AbsolutePosition
	End Property

	Public Property Get ActiveCommand()
		Set ActiveCommand = p_Resultset.ActiveCommand
	End Property

	Public Property Get ActiveConnection()
		If IsObject(p_Resultset.ActiveConnection) Then
			Set ActiveConnection = p_Resultset.ActiveConnection
		Else
			ActiveConnection = p_Resultset.ActiveConnection
		End If
	End Property

	Public Property Get BOF()
		BOF = p_Resultset.BOF
	End Property

	Public Property Get Bookmark()
		If IsObject(p_Resultset.Bookmark) Then
			Set Bookmark = p_Resultset.Bookmark
		Else
			Bookmark = p_Resultset.Bookmark
		End If 
	End Property

	Public Property Get CacheSize()
		CacheSize = p_Resultset.CacheSize
	End Property

	Public Property Get CursorLocation()
		Set CursorLocation = p_Resultset.CursorLocation
	End Property

	Public Property Get CursorType()
		Set CursorType = p_Resultset.CursorType
	End Property

	Public Property Get DataMember()
		DataMember = p_Resultset.DataMember
	End Property

	Public Property Get DataSource()
		If IsObject(p_Resultset.DataSource) Then
			Set DataSource = p_Resultset.DataSource
		Else
			DataSource = p_Resultset.DataSource
		End If 
	End Property

	Public Property Get EditMode()
		Set EditMode = p_Resultset.EditMode
	End Property

	Public Property Get EOF()
		EOF = p_Resultset.EOF
	End Property

	Public Property Get Fields()
		Set Fields = p_Resultset.Fields
	End Property

	Public Property Get Filter()
		If IsObject(p_Resultset.Filter) Then
			Set Filter = p_Resultset.Filter
		Else
			Filter = p_Resultset.Filter
		End If 
	End Property

	Public Property Get Index()
		Index = p_Resultset.Index
	End Property

	Public Property Get LockType()
		Set LockType = p_Resultset.LockType
	End Property

	Public Property Get MarshalOptions()
		Set MarshalOptions = p_Resultset.MarshalOptions
	End Property

	Public Property Get MaxRecords()
		MaxRecords = p_Resultset.MaxRecords
	End Property

	Public Property Get PageCount()
		PageCount = p_Resultset.PageCount
	End Property

	Public Property Get PageSize()
		PageSize = p_Resultset.PageSize
	End Property

	Public Property Get Properties()
		Set Properties = p_Resultset.Properties
	End Property

	Public Property Get RecordCount()
		RecordCount = p_Resultset.RecordCount
	End Property

	Public Property Get Sort()
		Sort = p_Resultset.Sort
	End Property

	Public Property Get Source()
		If IsObject(p_Resultset.Source) Then
			Set Source = p_Resultset.Source
		Else
			Source = p_Resultset.Source
		End If 
	End Property

	Public Property Get State()
		State = p_Resultset.State
	End Property

	Public Property Get Status()
		Status = p_Resultset.Status
	End Property

	Public Property Get StayInSync()
		StayInSync = p_Resultset.StayInSync
	End Property


	' Methods


	' Public Sub AddNew([FieldList], [Values])

	' End Sub

	Public Sub Cancel()
		p_Resultset.Cancel()
	End Sub

	' Public Sub CancelBatch([AffectRecords As AffectEnum = adAffectAll])

	' End Sub

	Public Sub CancelUpdate()
		p_Resultset.CancelUpdate()
	End Sub

	' Public Function Clone([LockType As LockTypeEnum = adLockUnspecified]) 

	' End Function

	Public Sub Close()
		p_Resultset.Close()
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
		p_Resultset.MoveFirst()
	End Sub

	Public Sub MoveLast()
		p_Resultset.MoveLast()
	End Sub

	Public Sub MoveNext()
		p_Resultset.MoveNext()
	End Sub

	Public Sub MovePrevious()
		p_Resultset.MovePrevious()
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
		Set p_Resultset = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_DB_Recordset.vbs" Then
	Dim resultset
	Set resultset = New base_DB_Resultset


End If
