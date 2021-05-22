Option Explicit

Class base_Outlook_ViewControl
	Private p_OutlookViewControl

	Private Sub Class_Initialize()
		Set p_OutlookViewControl = CreateObject("OVCtl.OVCtl")
	End Sub


	' Properties


	Public Property Get ActiveFolder()
		Set ActiveFolder = p_OutlookViewControl.ActiveFolder
	End Property

	Public Property Get DeferUpdate()
		DeferUpdate = p_OutlookViewControl.DeferUpdate 
	End Property

	Public Property Let DeferUpdate(blnDeferUpdate)
		p_OutlookViewControl.DeferUpdate = blnDeferUpdate
	End Property

	Public Property Get EnableRowPersistance()
		EnableRowPersistance = p_OutlookViewControl.EnableRowPersistance 
	End Property

	Public Property Let EnableRowPersistance(blnEnableRowPersistance)
		p_OutlookViewControl.EnableRowPersistance = blnEnableRowPersistance 
	End Property

	Public Property Get Filter()
		Filter = p_OutlookViewControl.Filter 
	End Property

	Public Property Let Filter(strFilter)
		p_OutlookViewControl.Filter = strFilter
	End Property

	Public Property Get FilterAppend()
		FilterAppend = p_OutlookViewControl.FilterAppend 
	End Property

	Public Property Let FilterAppend(strFilterAppend)
		p_OutlookViewControl.FilterAppend = strFilterAppend 
	End Property

	Public Property Get Folder()
		Folder = p_OutlookViewControl.Folder 
	End Property

	Public Property Let Folder(strFolder)
		p_OutlookViewControl.Folder = strFolder 
	End Property

	Public Property Get ItemCount()
		ItemCount = p_OutlookViewControl.ItemCount 
	End Property

	Public Property Get Namespace()
		Namespace = p_OutlookViewControl.Namespace 
	End Property

	Public Property Let Namespace(strNamespace)
		p_OutlookViewControl.Namespace = strNamespace
	End Property

	Public Property Get OutlookApplication()
		Set OutlookApplication = p_OutlookViewControl.OutlookApplication 
	End Property

	Public Property Get Restriction()
		Restriction = p_OutlookViewControl.Restriction 
	End Property

	Public Property Let Restriction(strRestriction)
		p_OutlookViewControl.Restriction = strRestriction
	End Property

	Public Property Get SelectedDate()
		SelectedDate = p_OutlookViewControl.SelectedDate 
	End Property

	Public Property Get Selection()
		Set Selection = p_OutlookViewControl.Selection 
	End Property

	Public Property Get View()
		View = p_OutlookViewControl.View 
	End Property

	Public Property Let View(strView)
		p_OutlookViewControl.View = strView 
	End Property

	Public Property Get ViewXML()
		ViewXML = p_OutlookViewControl.ViewXML 
	End Property

	Public Property Let ViewXML(strViewXML)
		p_OutlookViewControl.ViewXML = strViewXML
	End Property


	' Methods


	Public Sub AddressBook()
		p_OutlookViewControl.AddressBook
	End Sub

	Public Sub AddToPFFavorites()
		p_OutlookViewControl.AddToPFFavorites
	End Sub

	Public Sub AdvancedFind()
		p_OutlookViewControl.AdvancedFind
	End Sub

	Public Sub Categories()
		p_OutlookViewControl.Categories
	End Sub

	Public Sub CollapseAllGroups()
		p_OutlookViewControl.CollapseAllGroups
	End Sub

	Public Sub CollapseGroup()
		p_OutlookViewControl.CollapseGroup
	End Sub

	Public Sub CustomizeView()
		p_OutlookViewControl.CustomizeView
	End Sub

	Public Sub Delete()
		p_OutlookViewControl.Delete
	End Sub

	Public Sub ExpandAllGroups()
		p_OutlookViewControl.ExpandAllGroups
	End Sub

	Public Sub ExpandGroup()
		p_OutlookViewControl.ExpandGroup
	End Sub

	Public Sub FlagItem()
		p_OutlookViewControl.FlagItem
	End Sub

	Public Sub ForceUpdate()
		p_OutlookViewControl.ForceUpdate
	End Sub

	Public Sub Forward()
		p_OutlookViewControl.Forward
	End Sub

	Public Sub GoToDate(strNewDate)
		p_OutlookViewControl.GoToDate strNewDate
	End Sub

	Public Sub GoToToday()
		p_OutlookViewControl.GoToToday
	End Sub

	Public Sub GroupBy()
		p_OutlookViewControl.GroupBy
	End Sub

	Public Sub MarkAllAsRead()
		p_OutlookViewControl.MarkAllAsRead
	End Sub

	Public Sub MarkAsRead()
		p_OutlookViewControl.MarkAsRead
	End Sub

	Public Sub MarkAsUnread()
		p_OutlookViewControl.MarkAsUnread
	End Sub

	Public Sub MoveItem()
		p_OutlookViewControl.MoveItem
	End Sub

	Public Sub NewAppointment()
		p_OutlookViewControl.NewAppointment
	End Sub

	Public Sub NewContact()
		p_OutlookViewControl.NewContact
	End Sub

	Public Sub NewDefaultItem()
		p_OutlookViewControl.NewDefaultItem
	End Sub

	Public Sub NewDistributionList()
		p_OutlookViewControl.NewDistributionList
	End Sub

	Public Sub NewForm()
		p_OutlookViewControl.NewForm
	End Sub

	Public Sub NewJournalEntry()
		p_OutlookViewControl.NewJournalEntry
	End Sub

	Public Sub NewMeetingRequest()
		p_OutlookViewControl.NewMeetingRequest
	End Sub

	Public Sub NewMessage()
		p_OutlookViewControl.NewMessage
	End Sub

	Public Sub NewNote()
		p_OutlookViewControl.NewNote
	End Sub

	Public Sub NewPost()
		p_OutlookViewControl.NewPost
	End Sub

	Public Sub NewTask()
		p_OutlookViewControl.NewTask
	End Sub

	Public Sub NewTaskRequest()
		p_OutlookViewControl.NewTaskRequest
	End Sub

	Public Sub Open()
		p_OutlookViewControl.Open
	End Sub

	Public Sub OpenSharedDefaultFolder(strRecipient, objFolderType)
		p_OutlookViewControl.OpenSharedDefaultFolder strRecipient, objFolderType
	End Sub

	Public Sub PrintItem()
		p_OutlookViewControl.PrintItem
	End Sub

	Public Sub Reply()
		p_OutlookViewControl.Reply
	End Sub

	Public Sub ReplyAll()
		p_OutlookViewControl.ReplyAll
	End Sub

	Public Sub ReplyInFolder()
		p_OutlookViewControl.ReplyInFolder
	End Sub

	Public Sub SaveAs()
		p_OutlookViewControl.SaveAs
	End Sub

	Public Sub SendAndReceive()
		p_OutlookViewControl.SendAndReceive
	End Sub

	Public Sub ShowFields()
		p_OutlookViewControl.ShowFields
	End Sub

	Public Sub Sort()
		p_OutlookViewControl.Sort
	End Sub

	Public Sub SynchFolder()
		p_OutlookViewControl.SynchFolder
	End Sub


	' Events


	' Activate
	' BeforeViewSwitch
	' SelectionChange
	' ViewSwitch


	Private Sub Class_Terminate()
		Set p_OutlookViewControl = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Outlook_ViewControl.vbs" Then
	Dim objOVC
	Set objOVC = new base_Outlook_ViewControl

	objOVC.AddressBook
End If