Option Explicit

Class base_Word
	Private p_Word

	Private Sub Class_Initialize()
		Set p_Word = CreateObject("Word.Application")
	End Sub


	' Properties


	Public Property Get ActiveDocument()
		Set ActiveDocument = p_Word.ActiveDocument 
	End Property

	Public Property Get ActiveEncryptionSession()
		ActiveEncryptionSession = p_Word.ActiveEncryptionSession 
	End Property

	Public Property Get ActivePrinter()
		ActivePrinter = p_Word.ActivePrinter 
	End Property

	Public Property Let ActivePrinter(strActivePrinter)
		p_Word.ActivePrinter = strActivePrinter
	End Property

	Public Property Get ActiveProtectedViewWindow()
		Set ActiveProtectedViewWindow = p_Word.ActiveProtectedViewWindow 
	End Property

	Public Property Get ActiveWindow()
		Set ActiveWindow = p_Word.ActiveWindow 
	End Property

	Public Property Get AddIns()
		Set AddIns = p_Word.AddIns 
	End Property

	Public Property Get Application()
		Set Application = p_Word.Application 
	End Property

	Public Property Get ArbitraryXMLSupportAvailable()
		ArbitraryXMLSupportAvailable = p_Word.ArbitraryXMLSupportAvailable 
	End Property

	Public Property Get Assistance()
		Set Assistance = p_Word.Assistance 
	End Property

	Public Property Get AutoCaptions()
		Set AutoCaptions = p_Word.AutoCaptions 
	End Property

	Public Property Get AutoCorrect()
		Set AutoCorrect = p_Word.AutoCorrect 
	End Property

	Public Property Get AutoCorrectEmail()
		Set AutoCorrectEmail = p_Word.AutoCorrectEmail 
	End Property

	Public Property Get AutomationSecurity()
		Set AutomationSecurity = p_Word.AutomationSecurity 
	End Property

	Public Property Set AutomationSecurity(objMsoAutomationSecurity)
		Set p_Word.AutomationSecurity = objMsoAutomationSecurity
	End Property

	Public Property Get BackgroundPrintingStatus()
		BackgroundPrintingStatus = p_Word.BackgroundPrintingStatus 
	End Property

	Public Property Get BackgroundSavingStatus()
		BackgroundSavingStatus = p_Word.BackgroundSavingStatus 
	End Property

	Public Property Get Bibliography()
		Set Bibliography = p_Word.Bibliography 
	End Property

	Public Property Get BrowseExtraFileTypes()
		BrowseExtraFileTypes = p_Word.BrowseExtraFileTypes 
	End Property

	Public Property Let BrowseExtraFileTypes(strBrowseExtraFileTypes)
		p_Word.BrowseExtraFileTypes = strBrowseExtraFileTypes
	End Property

	Public Property Get Browser()
		Set Browser = p_Word.Browser 
	End Property

	Public Property Get Build()
		Build = p_Word.Build 
	End Property

	Public Property Get CapsLock()
		CapsLock = p_Word.CapsLock 
	End Property

	Public Property Get Caption()
		Caption = p_Word.Caption 
	End Property

	Public Property Let Caption(strCaption)
		p_Word.Caption = strCaption
	End Property

	Public Property Get CaptionLabels()
		Set CaptionLabels = p_Word.CaptionLabels 
	End Property

	Public Property Get ChartDataPointTrack()
		ChartDataPointTrack = p_Word.ChartDataPointTrack 
	End Property

	Public Property Let ChartDataPointTrack(blnChartDataPointTrack)
		p_Word.ChartDataPointTrack = blnChartDataPointTrack
	End Property

	Public Property Get CheckLanguage()
		CheckLanguage = p_Word.CheckLanguage 
	End Property

	Public Property Let CheckLanguage(blnCheckLanguage)
		p_Word.CheckLanguage = blnCheckLanguage
	End Property

	Public Property Get COMAddIns()
		Set COMAddIns = p_Word.COMAddIns 
	End Property

	Public Property Get CommandBars()
		Set CommandBars = p_Word.CommandBars 
	End Property

	Public Property Get Creator()
		Creator = p_Word.Creator 
	End Property

	Public Property Get CustomDictionaries()
		Set CustomDictionaries = p_Word.CustomDictionaries 
	End Property

	Public Property Get CustomizationContext()
		Set CustomizationContext = p_Word.CustomizationContext 
	End Property

	Public Property Set CustomizationContext(objCustomizationContext)
		Set p_Word.CustomizationContext = objCustomizationContext
	End Property

	Public Property Get DefaultLegalBlackline()
		DefaultLegalBlackline = p_Word.DefaultLegalBlackline 
	End Property

	Public Property Let DefaultLegalBlackline(blnDefaultLegalBlackline)
		p_Word.DefaultLegalBlackline = blnDefaultLegalBlackline
	End Property

	Public Property Get DefaultSaveFormat()
		DefaultSaveFormat = p_Word.DefaultSaveFormat 
	End Property

	Public Property Let DefaultSaveFormat(strDefaultSaveFormat)
		p_Word.DefaultSaveFormat = strDefaultSaveFormat
	End Property

	Public Property Get DefaultTableSeparator()
		DefaultTableSeparator = p_Word.DefaultTableSeparator 
	End Property

	Public Property Let DefaultTableSeparator(strDefaultTableSeparator)
		p_Word.DefaultTableSeparator = strDefaultTableSeparator
	End Property

	Public Property Get Dialogs()
		Set Dialogs = p_Word.Dialogs 
	End Property

	Public Property Get DisplayAlerts()
		Set DisplayAlerts = p_Word.DisplayAlerts 
	End Property

	Public Property Set DisplayAlerts(objWdAlertLevel)
		Set p_Word.DisplayAlerts = objWdAlertLevel
	End Property

	Public Property Get DisplayAutoCompleteTips()
		DisplayAutoCompleteTips = p_Word.DisplayAutoCompleteTips 
	End Property

	Public Property Let DisplayAutoCompleteTips(blnDisplayAutoCompleteTips)
		p_Word.DisplayAutoCompleteTips = blnDisplayAutoCompleteTips
	End Property

	Public Property Get DisplayDocumentInformationPanel()
		DisplayDocumentInformationPanel = p_Word.DisplayDocumentInformationPanel 
	End Property

	Public Property Let DisplayDocumentInformationPanel(blnDisplayDocumentInformationPanel)
		p_Word.DisplayDocumentInformationPanel = blnDisplayDocumentInformationPanel
	End Property

	Public Property Get DisplayRecentFiles()
		DisplayRecentFiles = p_Word.DisplayRecentFiles 
	End Property

	Public Property Let DisplayRecentFiles(blnDisplayRecentFiles)
		p_Word.DisplayRecentFiles = blnDisplayRecentFiles
	End Property

	Public Property Get DisplayScreenTips()
		DisplayScreenTips = p_Word.DisplayScreenTips 
	End Property

	Public Property Let DisplayScreenTips(blnDisplayScreenTips)
		p_Word.DisplayScreenTips = blnDisplayScreenTips
	End Property

	Public Property Get DisplayScrollBars()
		DisplayScrollBars = p_Word.DisplayScrollBars 
	End Property

	Public Property Let DisplayScrollBars(blnDisplayScrollBars)
		p_Word.DisplayScrollBars = blnDisplayScrollBars
	End Property

	Public Property Get Documents()
		Set Documents = p_Word.Documents 
	End Property

	Public Property Get DontResetInsertionPointProperties()
		DontResetInsertionPointProperties = p_Word.DontResetInsertionPointProperties 
	End Property

	Public Property Let DontResetInsertionPointProperties(blnDontResetInsertionPointProperties)
		p_Word.DontResetInsertionPointProperties = blnDontResetInsertionPointProperties
	End Property

	Public Property Get EmailOptions()
		Set EmailOptions = p_Word.EmailOptions 
	End Property

	Public Property Get EmailTemplate()
		EmailTemplate = p_Word.EmailTemplate 
	End Property

	Public Property Let EmailTemplate(strEmailTemplate)
		p_Word.EmailTemplate = strEmailTemplate
	End Property

	Public Property Get EnableCancelKey()
		Set EnableCancelKey = p_Word.EnableCancelKey 
	End Property

	Public Property Set EnableCancelKey(objWdEnableCancelKey)
		Set p_Word.EnableCancelKey = objWdEnableCancelKey
	End Property

	Public Property Get FeatureInstall()
		Set FeatureInstall = p_Word.FeatureInstall 
	End Property

	Public Property Set FeatureInstall(objMsoFeatureInstall)
		Set p_Word.FeatureInstall = objMsoFeatureInstall
	End Property

	Public Property Get FileConverters()
		Set FileConverters = p_Word.FileConverters 
	End Property

	Public Property Get FileDialog(objFileDialogType)
		Set FileDialog = p_Word.FileDialog(objFileDialogType)
	End Property

	Public Property Get FileValidation()
		Set FileValidation = p_Word.FileValidation 
	End Property

	Public Property Set FileValidation(objMsoFileValidationMode)
		Set p_Word.FileValidation = objMsoFileValidationMode
	End Property

	Public Property Get FindKey(lngKeyCode) ' Optional param: [KeyCode2]
		Set FindKey = p_Word.FindKey(lngKeyCode)
	End Property

	Public Property Get FocusInMailHeader()
		FocusInMailHeader = p_Word.FocusInMailHeader 
	End Property

	Public Property Get FontNames()
		FontNames = p_Word.FontNames 
	End Property

	Public Property Get HangulHanjaDictionaries()
		Set HangulHanjaDictionaries = p_Word.HangulHanjaDictionaries 
	End Property

	Public Property Get Height()
		Height = p_Word.Height 
	End Property

	Public Property Let Height(lngHeight)
		p_Word.Height = lngHeight
	End Property

	Public Property Get International(objInternationalIndex)
		Set International = p_Word.International(objInternationalIndex)
	End Property

	Public Property Get IsObjectValid(objObject)
		IsObjectValid = p_Word.IsObjectValid(objObject)
	End Property

	Public Property Get IsSandboxed()
		IsSandboxed = p_Word.IsSandboxed 
	End Property

	Public Property Get KeyBindings()
		Set KeyBindings = p_Word.KeyBindings 
	End Property

	Public Property Get KeysBoundTo(objKeyCategory, strCommand) ' Optional params: [CommandParameter]
		Set KeysBoundTo = p_Word.KeysBoundTo(objKeyCategory, strCommand)
	End Property

	Public Property Get LandscapeFontNames()
		Set LandscapeFontNames = p_Word.LandscapeFontNames 
	End Property

	Public Property Get Language()
		Set Language = p_Word.Language 
	End Property

	Public Property Get Languages()
		Set Languages = p_Word.Languages 
	End Property

	Public Property Get LanguageSettings()
		Set LanguageSettings = p_Word.LanguageSettings 
	End Property

	Public Property Get Left()
		Left = p_Word.Left 
	End Property

	Public Property Let Left(lngLeft)
		p_Word.Left = lngLeft
	End Property

	Public Property Get ListGalleries()
		Set ListGalleries = p_Word.ListGalleries 
	End Property

	Public Property Get MacroContainer()
		Set MacroContainer = p_Word.MacroContainer 
	End Property

	Public Property Get MailingLabel()
		Set MailingLabel = p_Word.MailingLabel 
	End Property

	Public Property Get MailMessage()
		Set MailMessage = p_Word.MailMessage 
	End Property

	Public Property Get MailSystem()
		Set MailSystem = p_Word.MailSystem 
	End Property

	Public Property Get MAPIAvailable()
		MAPIAvailable = p_Word.MAPIAvailable 
	End Property

	Public Property Get MathCoprocessorAvailable()
		MathCoprocessorAvailable = p_Word.MathCoprocessorAvailable 
	End Property

	Public Property Get MouseAvailable()
		MouseAvailable = p_Word.MouseAvailable 
	End Property

	Public Property Get Name()
		Name = p_Word.Name 
	End Property

	Public Property Get NewDocument()
		Set NewDocument = p_Word.NewDocument 
	End Property

	Public Property Get NormalTemplate()
		Set NormalTemplate = p_Word.NormalTemplate 
	End Property

	Public Property Get NumLock()
		NumLock = p_Word.NumLock 
	End Property

	Public Property Get OMathAutoCorrect()
		Set OMathAutoCorrect = p_Word.OMathAutoCorrect 
	End Property

	Public Property Get OpenAttachmentsInFullScreen()
		OpenAttachmentsInFullScreen = p_Word.OpenAttachmentsInFullScreen 
	End Property

	Public Property Let OpenAttachmentsInFullScreen(blnOpenAttachmentsInFullScreen)
		p_Word.OpenAttachmentsInFullScreen = blnOpenAttachmentsInFullScreen
	End Property

	Public Property Get Options()
		Set Options = p_Word.Options 
	End Property

	Public Property Get Parent()
		Set Parent = p_Word.Parent 
	End Property

	Public Property Get Path()
		Path = p_Word.Path 
	End Property

	Public Property Get PathSeparator()
		PathSeparator = p_Word.PathSeparator 
	End Property

	Public Property Get PickerDialog()
		Set PickerDialog = p_Word.PickerDialog 
	End Property

	Public Property Get PortraitFontNames()
		Set PortraitFontNames = p_Word.PortraitFontNames 
	End Property

	Public Property Get PrintPreview()
		PrintPreview = p_Word.PrintPreview 
	End Property

	Public Property Let PrintPreview(blnPrintPreview)
		p_Word.PrintPreview = blnPrintPreview
	End Property

	Public Property Get ProtectedViewWindows()
		Set ProtectedViewWindows = p_Word.ProtectedViewWindows 
	End Property

	Public Property Get RecentFiles()
		Set RecentFiles = p_Word.RecentFiles 
	End Property

	Public Property Get RestrictLinkedStyles()
		RestrictLinkedStyles = p_Word.RestrictLinkedStyles 
	End Property

	Public Property Let RestrictLinkedStyles(blnRestrictLinkedStyles)
		p_Word.RestrictLinkedStyles = blnRestrictLinkedStyles
	End Property

	Public Property Get ScreenUpdating()
		ScreenUpdating = p_Word.ScreenUpdating 
	End Property

	Public Property Let ScreenUpdating(blnScreenUpdating)
		p_Word.ScreenUpdating = blnScreenUpdating
	End Property

	Public Property Get Selection()
		Set Selection = p_Word.Selection 
	End Property

	Public Property Get ShowAnimation()
		ShowAnimation = p_Word.ShowAnimation 
	End Property

	Public Property Let ShowAnimation(blnShowAnimation)
		p_Word.ShowAnimation = blnShowAnimation
	End Property

	Public Property Get ShowStartupDialog()
		ShowStartupDialog = p_Word.ShowStartupDialog 
	End Property

	Public Property Let ShowStartupDialog(blnShowStartupDialog)
		p_Word.ShowStartupDialog = blnShowStartupDialog
	End Property

	Public Property Get ShowStylePreviews()
		ShowStylePreviews = p_Word.ShowStylePreviews 
	End Property

	Public Property Let ShowStylePreviews(blnShowStylePreviews)
		p_Word.ShowStylePreviews = blnShowStylePreviews
	End Property

	Public Property Get ShowVisualBasicEditor()
		ShowVisualBasicEditor = p_Word.ShowVisualBasicEditor 
	End Property

	Public Property Let ShowVisualBasicEditor(blnShowVisualBasicEditor)
		p_Word.ShowVisualBasicEditor = blnShowVisualBasicEditor
	End Property

	Public Property Get SmartArtColors()
		Set SmartArtColors = p_Word.SmartArtColors 
	End Property

	Public Property Get SmartArtLayouts()
		Set SmartArtLayouts = p_Word.SmartArtLayouts 
	End Property

	Public Property Get SmartArtQuickStyles()
		Set SmartArtQuickStyles = p_Word.SmartArtQuickStyles 
	End Property

	Public Property Get SpecialMode()
		SpecialMode = p_Word.SpecialMode 
	End Property

	Public Property Get StartupPath()
		StartupPath = p_Word.StartupPath 
	End Property

	Public Property Let StartupPath(strPath)
		p_Word.StartupPath = strPath
	End Property

	Public Property Get StatusBar()
		StatusBar = p_Word.StatusBar 
	End Property

	Public Property Let StatusBar(strStatusBar)
		p_Word.StatusBar = strStatusBar
	End Property

	Public Property Get SynonymInfo(strWord) ' Optional params: [LanguageID]
		Set SynonymInfo = p_Word.SynonymInfo(strWord)
	End Property

	Public Property Get System()
		Set System = p_Word.System 
	End Property

	Public Property Get TaskPanes()
		Set TaskPanes = p_Word.TaskPanes 
	End Property

	Public Property Get Tasks()
		Set Tasks = p_Word.Tasks 
	End Property

	Public Property Get Templates()
		Set Templates = p_Word.Templates 
	End Property

	Public Property Get Top()
		Top = p_Word.Top 
	End Property

	Public Property Let Top(lngTop)
		p_Word.Top = lngTop
	End Property

	Public Property Get UndoRecord()
		Set UndoRecord = p_Word.UndoRecord 
	End Property

	Public Property Get UsableHeight()
		UsableHeight = p_Word.UsableHeight 
	End Property

	Public Property Get UsableWidth()
		UsableWidth = p_Word.UsableWidth 
	End Property

	Public Property Get UserAddress()
		UserAddress = p_Word.UserAddress 
	End Property

	Public Property Let UserAddress(strUserAddress)
		p_Word.UserAddress = strUserAddress
	End Property

	Public Property Get UserControl()
		UserControl = p_Word.UserControl 
	End Property

	Public Property Get UserInitials()
		UserInitials = p_Word.UserInitials 
	End Property

	Public Property Let UserInitials(strUserInitials)
		p_Word.UserInitials = strUserInitials
	End Property

	Public Property Get UserName()
		UserName = p_Word.UserName 
	End Property

	Public Property Let UserName(strUserName)
		p_Word.UserName = strUserName
	End Property

	Public Property Get VBE()
		Set VBE = p_Word.VBE 
	End Property

	Public Property Get Version()
		Version = p_Word.Version 
	End Property

	Public Property Get Visible()
		Visible = p_Word.Visible 
	End Property

	Public Property Let Visible(blnVisible)
		p_Word.Visible = blnVisible
	End Property

	Public Property Get Width()
		Width = p_Word.Width 
	End Property

	Public Property Let Width(lngWidth)
		p_Word.Width = lngWidth
	End Property

	Public Property Get Windows()
		Set Windows = p_Word.Windows 
	End Property

	Public Property Get WindowState()
		Set WindowState = p_Word.WindowState 
	End Property

	Public Property Let WindowState(objWdWindowState)
		Set p_Word.WindowState = objWdWindowState
	End Property

	Public Property Get WordBasic()
		Set WordBasic = p_Word.WordBasic 
	End Property

	Public Property Get XMLNamespaces()
		Set XMLNamespaces = p_Word.XMLNamespaces 
	End Property



	' Methods


	Public Sub Activate()
		p_Word.Activate
	End Sub

	Public Sub AddAddress(strTagID(), strValue())
		p_Word.AddAddress strTagID(), strValue()
	End Sub

	Public Sub AutomaticChange()
		p_Word.AutomaticChange
	End Sub

	Public Function BuildKeyCode(objArg1) ' Optional params: [Arg2], [Arg3], [Arg4]) As Long
		BuildKeyCode = p_Word.BuildKeyCode(objArg1)
	End Function

	Public Function CentimetersToPoints(sngCentimeters)
		CentimetersToPoints = p_Word.CentimetersToPoints(sngCentimeters)
	End Function

	Public Sub ChangeFileOpenDirectory(strPath)
		p_Word.ChangeFileOpenDirectory strPath
	End Sub

	Public Function CheckGrammar(strString)
		CheckGrammar = p_Word.CheckGrammar(strString)
	End Function

	Public Function CheckSpelling(strWord) ' Optional params: [CustomDictionary], [IgnoreUppercase], [MainDictionary], [CustomDictionary2], [CustomDictionary3], [CustomDictionary4], [CustomDictionary5], [CustomDictionary6], [CustomDictionary7], [CustomDictionary8], [CustomDictionary9], [CustomDictionary10]) As Boolean
		CheckSpelling = p_Word.CheckSpelling(strWord)
	End Function

	Public Function CleanString(strString)
		CleanString = p_Word.CleanString(strString)
	End Function

	Public Function CompareDocuments(objOriginalDocument, objRevisedDocument) ' Optional params: [Destination As WdCompareDestination = wdCompareDestinationNew], [Granularity As WdGranularity = wdGranularityWordLevel], [CompareFormatting As Boolean = True], [CompareCaseChanges As Boolean = True], [CompareWhitespace As Boolean = True], [CompareTables As Boolean = True], [CompareHeaders As Boolean = True], [CompareFootnotes As Boolean = True], [CompareTextboxes As Boolean = True], [CompareFields As Boolean = True], [CompareComments As Boolean = True], [CompareMoves As Boolean = True], [RevisedAuthor As String], [IgnoreAllComparisonWarnings As Boolean = False]) As Document
		Set CompareDocuments = p_Word.CompareDocuments(objOriginalDocument, objRevisedDocument)
	End Function

	Public Sub DDEExecute(lngChannel, strCommand)
		p_Word.DDEExecute lngChannel, strCommand
	End Sub

	Public Function DDEInitiate(strApp, strTopic)
		DDEInitiate = p_Word.DDEInitiate(strApp, strTopic)
	End Function

	Public Sub DDEPoke(lngChannel, strItem, strData)
		p_Word.DDEPoke lngChannel, strItem, strData
	End Sub

	Public Function DDERequest(lngChannel, strItem)
		DDERequest = p_Word.DDERequest(lngChannel, strItem)
	End Function

	Public Sub DDETerminate(lngChannel)
		p_Word.DDETerminate lngChannel
	End Sub

	Public Sub DDETerminateAll()
		p_Word.DDETerminateAll
	End Sub

	Public Function DefaultWebOptions()
		Set DefaultWebOptions = p_Word.DefaultWebOptions
	End Function

	Public Function GetAddress() ' Optional params: [Name], [AddressProperties], [UseAutoText], [DisplaySelectDialog], [SelectDialog], [CheckNamesDialog], [RecentAddressesChoice], [UpdateRecentAddresses]) As String
		GetAddress = p_Word.GetAddress
	End Function

	Public Function GetDefaultTheme(objDocumentType)
		GetDefaultTheme = p_Word.GetDefaultTheme(objDocumentType)
	End Function

	Public Function GetSpellingSuggestions(strWord) ' Optional params: [CustomDictionary], [IgnoreUppercase], [MainDictionary], [SuggestionMode], [CustomDictionary2], [CustomDictionary3], [CustomDictionary4], [CustomDictionary5], [CustomDictionary6], [CustomDictionary7], [CustomDictionary8], [CustomDictionary9], [CustomDictionary10]) As SpellingSuggestions
		Set GetSpellingSuggestions = p_Word.GetSpellingSuggestions
	End Function

	Public Sub GoBack()
		p_Word.GoBack
	End Sub

	Public Sub GoForward()
		p_Word.GoForward
	End Sub

	Public Sub Help(varHelpType)
		p_Word.Help varHelpType
	End Sub

	Public Sub HelpTool()
		p_Word.HelpTool
	End Sub

	Public Function InchesToPoints(sngInches)
		InchesToPoints = p_Word.InchesToPoints(sngInches)
	End Function

	Public Function Keyboard() ' Optional params: [LangId As Long])
		Keyboard = p_Word.Keyboard()
	End Function

	Public Sub KeyboardBidi()
		p_Word.KeyboardBidi
	End Sub

	Public Sub KeyboardLatin()
		p_Word.KeyboardLatin
	End Sub

	Public Function KeyString(lngKeyCode) ' Optional params: [KeyCode2]
		KeyString = p_Word.KeyString(lngKeyCode)
	End Function

	Public Function LinesToPoints(sngLines)
		LinesToPoints = p_Word.LinesToPoints(sngLines)
	End Function

	Public Sub ListCommands(blnListAllCommands)
		p_Word.ListCommands blnListAllCommands
	End Sub

	Public Sub LoadMasterList(strFileName)
		p_Word.LoadMasterList strFileName
	End Sub

	Public Sub LookupNameProperties(strName)
		p_Word.LookupNameProperties strName
	End Sub

	Public Function MergeDocuments(objOriginalDocument, objRevisedDocument) ' Optional params: [Destination As WdCompareDestination = wdCompareDestinationNew], [Granularity As WdGranularity = wdGranularityWordLevel], [CompareFormatting As Boolean = True], [CompareCaseChanges As Boolean = True], [CompareWhitespace As Boolean = True], [CompareTables As Boolean = True], [CompareHeaders As Boolean = True], [CompareFootnotes As Boolean = True], [CompareTextboxes As Boolean = True], [CompareFields As Boolean = True], [CompareComments As Boolean = True], [CompareMoves As Boolean = True], [OriginalAuthor As String], [RevisedAuthor As String], [FormatFrom As WdMergeFormatFrom = wdMergeFormatFromPrompt]) As Document
		Set MergeDocuments = p_Word.MergeDocuments(objOriginalDocument, objRevisedDocument)
	End Function

	Public Function MillimetersToPoints(sngMillimeters)
		MillimetersToPoints = p_Word.MillimetersToPoints(sngMillimeters)
	End Function

	Public Sub Move(lngLeft, lngTop)
		p_Word.Move lngLeft, lngTop
	End Sub

	Public Function NewWindow()
		Set NewWindow = p_Word.NewWindow()
	End Function

	Public Sub NextLetter()
		p_Word.NextLetter
	End Sub

	Public Sub OnTime(varWhen, strName) ' Optional param: [Tolerance]
		p_Word.OnTime varWhen, strName
	End Sub

	Public Sub OrganizerCopy(strSource, strDestination, strName, objWdOrganizerObject)
		p_Word.OrganizerCopy strSource, strDestination, strName, objWdOrganizerObject
	End Sub

	Public Sub OrganizerDelete(strSource, strName, objWdOrganizerObject)
		p_Word.OrganizerDelete strSource, strName, objWdOrganizerObject
	End Sub

	Public Sub OrganizerRename(strSource, strName, strNewName, objWdOrganizerObject)
		p_Word.OrganizerRename strSource, strName, strNewName, objWdOrganizerObject
	End Sub

	Public Function PicasToPoints(sngPicas)
		PicasToPoints = p_Word.PicasToPoints(sngPicas)
	End Function

	Public Function PixelsToPoints(sngPixels) ' Optional params: [fVertical])
		PixelsToPoints = p_Word.PixelsToPoints(sngPixels)
	End Function

	Public Function PointsToCentimeters(sngPoints)
		PointsToCentimeters = p_Word.PointsToCentimeters(sngPoints)
	End Function

	Public Function PointsToInches(sngPoints)
		PointsToInches = p_Word.PointsToInches(sngPoints)
	End Function

	Public Function PointsToLines(sngPoints)
		PointsToLines = p_Word.PointsToLines(sngPoints)
	End Function

	Public Function PointsToMillimeters(sngPoints)
		PointsToMillimeters = p_Word.PointsToMillimeters(sngPoints)
	End Function

	Public Function PointsToPicas(sngPoints)
		PointsToPicas = p_Word.PointsToPicas(sngPoints)
	End Function

	Public Function PointsToPixels(sngPixels) ' Optional params: [fVertical])
		PointsToPixels = p_Word.PointsToPixels(sngPoints)
	End Function

	Public Sub PrintOut() ' Optional params: [Background], [Append], [Range], [OutputFileName], [From], [To], [Item], [Copies], [Pages], [PageType], [PrintToFile], [Collate], [FileName], [ActivePrinterMacGX], [ManualDuplexPrint], [PrintZoomColumn], [PrintZoomRow], [PrintZoomPaperWidth], [PrintZoomPaperHeight])
		p_Word.PrintOut
	End Sub

	Public Function ProductCode()
		ProductCode = p_Word.ProductCode()
	End Function

	Public Sub Quit() ' Optional params: [SaveChanges], [OriginalFormat], [RouteDocument])
		p_Word.Quit
	End Sub

	Public Function Repeat() ' Optional params: [Times]
		Repeat = p_Word.Repeat()
	End Function

	Public Sub ResetIgnoreAll()
		p_Word.ResetIgnoreAll
	End Sub

	Public Sub Resize(lngWidth, lngHeight)
		p_Word.Resize lngWidth, lngHeight
	End Sub

	Public Function Run(strMacroName) ' Optional params: [varg1], [varg2], [varg3], [varg4], [varg5], [varg6], [varg7], [varg8], [varg9], [varg10], [varg11], [varg12], [varg13], [varg14], [varg15], [varg16], [varg17], [varg18], [varg19], [varg20], [varg21], [varg22], [varg23], [varg24], [varg25], [varg26], [varg27], [varg28], [varg29], [varg30])
		Run = p_Word.Run(strMacroName)
	End Function

	Public Sub ScreenRefresh()
		p_Word.ScreenRefresh
	End Sub

	Public Sub SetDefaultTheme(strName, objDocumentType)
		p_Word.SetDefaultTheme strName, objDocumentType
	End Sub

	Public Sub ShowClipboard()
		p_Word.ShowClipboard
	End Sub

	Public Sub ShowMe()
		p_Word.ShowMe
	End Sub

	Public Sub SubstituteFont(strUnavailableFont, strSubstituteFont)
		p_Word.SubstituteFont strUnavailableFont, strSubstituteFont
	End Sub

	Public Sub ToggleKeyboard()
		p_Word.ToggleKeyboard
	End Sub


	' Events


	' DocumentBeforeClose
	' DocumentBeforePrint
	' DocumentBeforeSave
	' DocumentChange
	' DocumentOpen
	' DocumentSync
	' EPostageInsert
	' EPostageInsertEx
	' EPostagePropertyDialog
	' MailMergeAfterMerge
	' MailMergeAfterRecordMerge
	' MailMergeBeforeMerge
	' MailMergeBeforeRecordMerge
	' MailMergeDataSourceLoad
	' MailMergeDataSourceValidate
	' MailMergeDataSourceValidate2
	' MailMergeWizardSendToCustom
	' MailMergeWizardStateChange
	' NewDocument
	' ProtectedViewWindowActivate
	' ProtectedViewWindowBeforeClose
	' ProtectedViewWindowBeforeEdit
	' ProtectedViewWindowDeactivate
	' ProtectedViewWindowOpen
	' ProtectedViewWindowSize
	' Quit
	' WindowActivate
	' WindowBeforeDoubleClick
	' WindowBeforeRightClick
	' WindowDeactivate
	' WindowSelectionChange
	' WindowSize
	' XMLSelectionChange
	' XMLValidationError


	Private Sub Class_Terminate()
		Set p_Word = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Word.vbs" Then

End If
