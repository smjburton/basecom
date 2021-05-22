Option Explicit

' Include "base_Database_Result"

Class base_Database_Resultset
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


	Public Sub AddNew() ' Optional params: [FieldList], [Values])
		p_Resultset.AddNew
	End Sub

	Public Sub Cancel()
		p_Resultset.Cancel
	End Sub

	Public Sub CancelBatch() ' Optional params: [AffectRecords As AffectEnum = adAffectAll]
		p_Resultset.CancelBatch
	End Sub

	Public Sub CancelUpdate()
		p_Resultset.CancelUpdate
	End Sub

	Public Function Clone() ' Optional params: [LockType As LockTypeEnum = adLockUnspecified]) 
		Set Clone = p_Resultset.Clone()
	End Function

	Public Sub Close()
		p_Resultset.Close
	End Sub

	Public Function CompareBookmarks(varBookmark1, varBookmark2)
		CompareBookmarks = p_Resultset.CompareBookmarks(varBookmark1, varBookmark2)
	End Function

	Public Sub Delete() ' Optional params: [AffectRecords As AffectEnum = adAffectCurrent])
		p_Resultset.Delete
	End Sub

	Public Sub Find(strCriteria) ' Optional params: [SkipRecords As Long], [SearchDirection As SearchDirectionEnum = adSearchForward], [Start])
		p_Resultset.Find strCriteria
	End Sub

	Public Function GetRows() ' Optional params: [Rows As Long = -1], [Start], [Fields])
		Set GetRows = p_Resultset.GetRows()
	End Function

	Public Function GetString() ' Optional params: [StringFormat As StringFormatEnum = adClipString], [NumRows As Long = -1], [ColumnDelimeter As String], [RowDelimeter As String], [NullExpr As String])
		GetString = p_Resultset.GetString()
	End Function
 
	Public Sub Move(lngNumberOfRecords) ' Optional params: [Start]
		p_Resultset.Move lngNumberOfRecords
	End Sub

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

	Public Function NextRecordset() ' Optional params: [RecordsAffected])
		Set NextRecordset = p_Resultset.NextRecordset()
	End Function

	Public Sub Open() ' Optional params: [Source], [ActiveConnection], [CursorType As CursorTypeEnum = adOpenUnspecified], [LockType As LockTypeEnum = adLockUnspecified], [Options As Long = -1])
		p_Resultset.Open
	End Sub

	Public Sub ReQuery() ' Optional params: [Options As Long = -1])
		p_Resultset.ReQuery
	End Sub

	Public Sub ReSync() ' Optional params: [AffectRecords As AffectEnum = adAffectAll], [ResyncValues As ResyncEnum = adResyncAllValues])
		p_Resultset.ReSync
	End Sub

	Public Sub Save() ' Optional params: [Destination], [PersistFormat As PersistFormatEnum = adPersistADTG])
		p_Resultset.Save
	End Sub

	Public Sub Seek(varKeyValues) ' Optional params: [SeekOption As SeekEnum = adSeekFirstEQ])
		p_Resultset.Seek varKeyValues
	End Sub

	Public Function Supports(intCursorOptionEnum)
		Supports = p_Resultset.Supports(intCursorOptionEnum)
	End Function
 
	Public Sub Update() ' Optional params: ' [Fields], [Values])
		p_Resultset.Update
	End Sub

	Public Sub UpdateBatch() ' Optional params: ' [AffectRecords As AffectEnum = adAffectAll])
		p_Resultset.UpdateBatch
	End Sub

	Private Sub Class_Terminate()
		Set p_Resultset = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_Resultset.vbs" Then
	Dim resultset
	Set resultset = New base_Database_Resultset
End If
