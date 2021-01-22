Option Explicit

Class base_Outlook
	Private p_Outlook

	Private Sub Class_Initialize()
		Set p_Outlook = CreateObject("Outlook.Application")
	End Sub


	' Properties


	Public Property Get Application()
		Set Application = p_Outlook.Application 
	End Property

	Public Property Get Assistance()
		Set Assistance  = p_Outlook.Assistance
	End Property

	Public Property Get OutlookClass()
		Set OutlookClass = p_Outlook.Class
	End Property

	Public Property Get ComAddIns()
		Set ComAddIns = p_Outlook.COMAddIns
	End Property

	Public Property Get DefaultProfileName()
		DefaultProfileName = p_Outlook.DefaultProfileName
	End Property

	Public Property Get Explorers()
		Set Explorers = p_Outlook.Explorers
	End Property

	Public Property Get Inspectors()
		Set Inspectors = p_Outlook.Inspectors
	End Property

	Public Property Get IsTrusted
		IsTrusted = p_Outlook.IsTrusted
	End Property

	Public Property Get LanguageSettings()
		Set LanguageSettings = p_Outlook.LanguageSettings
	End Property

	Public Property Get Name()
		Name = p_Outlook.Name
	End Property

	Public Property Get Parent()
		Set Parent = p_Outlook.Parent
	End Property


	' Methods


	Public Function ActiveExplorer()
		Set ActiveExplorer = p_Outlook.ActiveExplorer()
	End Function

	Public Function ActiveInspector()
		Set ActiveInspector = p_Outlook.ActiveInspector()
	End Function

	Public Function ActiveWindow()
		Set ActiveWindow = p_Outlook.ActiveWindow()
	End Function

	Public Function AdvancedSearch(strScope) ' Optional params: [Filter], [SearchSubFolders], [Tag]) As Search
		Set AdvancedSearch = p_Outlook.AdvancedSearch(strScope)
	End Function

	Public Function CopyFile(strFilePath, strDestFolderPath)
		Set CopyFile = p_Outlook.CopyFile(strFilePath, strDestFolderPath)
	End Function

	Public Function CreateItem(objItemType)
		Set CreateItem = p_Outlook.CreateItem(objItemType)
	End Function

	Public Function CreateItemFromTemplate(strTemplatePath) ' Optional params: [InFolder]) As Object
		Set CreateItemFromTemplate = p_Outlook.CreateItemFromTemplate(strTemplatePath)
	End Function

	Public Function CreateObject(strObjectName)
		Set CreateObject = p_Outlook.CreateObject(strObjectName)
	End Function

	Public Function GetNamespace(strType)
		Set GetNamespace = p_Outlook.GetNamespace(strType)
	End Function

	Public Function GetObjectReference(objItem, objReferenceType)
		Set GetObjectReference = p_Outlook.GetObjectReference(objItem, objReferenceType)
	End Function

	Public Function IsSearchSynchronous(strLookInFolders)
		IsSearchSynchronous = p_Outlook.IsSearchSynchronous(strLookInFolders)
	End Function

	Public Sub Quit()
		p_Outlook.Quit
	End Sub

	Public Sub RefreshFormRegionDefinition(strRegionName)
		p_Outlook.RefreshFormRegionDefinition(strRegionName)
	End Sub


	' Events


	' AdvancedSearchComplete
	' AdvancedSearchStopped
	' BeforeFolderSharingDialog
	' ItemLoad
	' ItemSend
	' MAPILogonComplete
	' NewMail
	' NewMailEx
	' OptionsPagesAdd
	' Quit
	' Reminder
	' Startup


	Private Sub Class_Terminate()
		Set p_Outlook = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Outlook.vbs" Then

End If