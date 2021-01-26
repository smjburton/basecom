Option Explicit

' Look at:
' http://pandas.pydata.org/pandas-docs/stable/api.html
' http://docs.xlwings.org/api.html
' https://openpyxl.readthedocs.org/en/default/
' https://openpyxl.readthedocs.org/en/default/api/openpyxl.html

' Constants

Const xlDelimited = 1
Const xlTextQualifierDoubleQuote = 1
Const xlToLeft = -4159
Const xlFilterValues = 7
Const xlCenter = -4108

Class base_Excel
	Private p_Excel


	' Initialization


	Private Sub Class_Initialize()
		Set p_Excel = CreateObject("Excel.Application")
	End Sub


	' Properties


	Public Property Get ActiveCell()
		ActiveCell = p_Excel.ActiveCell
	End Property

	Public Property Get ActiveChart()
		ActiveChart = p_Excel.ActiveChart
	End Property

	Public Property Get ActiveEncryptionSession()
		ActiveEncryptionSession = p_Excel.ActiveEncryptionSession
	End Property

	Public Property Get ActivePrinter()
		ActivePrinter = p_Excel.ActivePrinter
	End Property

	Public Property Get ActiveProtectedViewWindow()
		ActiveProtectedViewWindow = p_Excel.ActiveProtectedViewWindow
	End Property

	Public Property Get ActiveSheet()
		ActiveSheet = p_Excel.ActiveSheet
	End Property

	Public Property Get ActiveWindow()
		ActiveWindow = p_Excel.ActiveWindow
	End Property

	Public Property Get ActiveWorkbook()
		ActiveWorkbook = p_Excel.ActiveWorkbook
	End Property

	Public Property Get AddIns()
		AddIns = p_Excel.AddIns
	End Property

	Public Property Get AddIns2()
		AddIns2 = p_Excel.AddIns2
	End Property

	Public Property Get AlertBeforeOverwriting()
		AlertBeforeOverwriting = p_Excel.AlertBeforeOverwriting
	End Property

	Public Property Get AltStartupPath()
		AltStartupPath = p_Excel.AltStartupPath
	End Property

	Public Property Get AlwaysUseClearType()
		AlwaysUseClearType = p_Excel.AlwaysUseClearType
	End Property

	Public Property Get Application()
		Application = p_Excel.Application
	End Property

	Public Property Get ArbitraryXMLSupportAvailable()
		ArbitraryXMLSupportAvailable = p_Excel.ArbitraryXMLSupportAvailable
	End Property

	Public Property Get AskToUpdateLinks()
		AskToUpdateLinks = p_Excel.AskToUpdateLinks
	End Property

	Public Property Get Assistance()
		Assistance = p_Excel.Assistance
	End Property

	Public Property Get AutoCorrect()
		AutoCorrect = p_Excel.AutoCorrect
	End Property

	Public Property Get AutoFormatAsYouTypeReplaceHyperlinks()
		AutoFormatAsYouTypeReplaceHyperlinks = p_Excel.AutoFormatAsYouTypeReplaceHyperlinks
	End Property

	Public Property Get AutomationSecurity()
		AutomationSecurity = p_Excel.AutomationSecurity
	End Property

	Public Property Get AutoPercentEntry()
		AutoPercentEntry = p_Excel.AutoPercentEntry
	End Property

	Public Property Get AutoRecover()
		AutoRecover = p_Excel.AutoRecover
	End Property

	Public Property Get Build()
		Build = p_Excel.Build
	End Property

	Public Property Get CalculateBeforeSave()
		CalculateBeforeSave = p_Excel.CalculateBeforeSave
	End Property

	Public Property Get Calculation()
		Calculation = p_Excel.Calculation
	End Property

	Public Property Get CalculationInterruptKey()
		CalculationInterruptKey = p_Excel.CalculationInterruptKey
	End Property

	Public Property Get CalculationState()
		CalculationState = p_Excel.CalculationState
	End Property

	Public Property Get CalculationVersion()
		CalculationVersion = p_Excel.CalculationVersion
	End Property

	Public Property Get Caller()
		Caller = p_Excel.Caller
	End Property

	Public Property Get CanPlaySounds()
		CanPlaySounds = p_Excel.CanPlaySounds
	End Property

	Public Property Get CanRecordSounds()
		CanRecordSounds = p_Excel.CanRecordSounds
	End Property

	Public Property Get Caption()
		Caption = p_Excel.Caption
	End Property

	Public Property Get CellDragAndDrop()
		CellDragAndDrop = p_Excel.CellDragAndDrop
	End Property

	Public Property Get Cells()
		Cells = p_Excel.Cells
	End Property

	Public Property Get ChartDataPointTrack()
		ChartDataPointTrack = p_Excel.ChartDataPointTrack
	End Property

	Public Property Get Charts()
		Charts = p_Excel.Charts
	End Property

	Public Property Get ClipboardFormats()
		ClipboardFormats = p_Excel.ClipboardFormats
	End Property

	Public Property Get ClusterConnector()
		ClusterConnector = p_Excel.ClusterConnector
	End Property

	Public Property Get Columns()
		Columns = p_Excel.Columns
	End Property

	Public Property Get COMAddIns()
		COMAddIns = p_Excel.COMAddIns
	End Property

	Public Property Get CommandBars()
		CommandBars = p_Excel.CommandBars
	End Property

	Public Property Get CommandUnderlines()
		CommandUnderlines = p_Excel.CommandUnderlines
	End Property

	Public Property Get ConstrainNumeric()
		ConstrainNumeric = p_Excel.ConstrainNumeric
	End Property

	Public Property Get ControlCharacters()
		ControlCharacters = p_Excel.ControlCharacters
	End Property

	Public Property Get CopyObjectsWithCells()
		CopyObjectsWithCells = p_Excel.CopyObjectsWithCells
	End Property

	Public Property Get Creator()
		Creator = p_Excel.Creator
	End Property

	Public Property Get Cursor()
		Cursor = p_Excel.Cursor
	End Property

	Public Property Get CursorMovement()
		CursorMovement = p_Excel.CursorMovement
	End Property

	Public Property Get CustomListCount()
		CustomListCount = p_Excel.CustomListCount
	End Property

	Public Property Get CutCopyMode()
		CutCopyMode = p_Excel.CutCopyMode
	End Property

	Public Property Get DataEntryMode()
		DataEntryMode = p_Excel.DataEntryMode
	End Property

	Public Property Get DDEAppReturnCode()
		DDEAppReturnCode = p_Excel.DDEAppReturnCode
	End Property

	Public Property Get DecimalSeparator()
		DecimalSeparator = p_Excel.DecimalSeparator
	End Property

	Public Property Get DefaultFilePath()
		DefaultFilePath = p_Excel.DefaultFilePath
	End Property

	Public Property Get DefaultSaveFormat()
		DefaultSaveFormat = p_Excel.DefaultSaveFormat
	End Property

	Public Property Get DefaultSheetDirection()
		DefaultSheetDirection = p_Excel.DefaultSheetDirection
	End Property

	Public Property Get DefaultWebOptions()
		DefaultWebOptions = p_Excel.DefaultWebOptions
	End Property

	Public Property Get DeferAsyncQueries()
		DeferAsyncQueries = p_Excel.DeferAsyncQueries
	End Property

	Public Property Get Dialogs()
		Dialogs = p_Excel.Dialogs
	End Property

	Public Property Get DisplayAlerts()
		DisplayAlerts = p_Excel.DisplayAlerts
	End Property

	Public Property Get DisplayClipboardWindow()
		DisplayClipboardWindow = p_Excel.DisplayClipboardWindow
	End Property

	Public Property Get DisplayCommentIndicator()
		DisplayCommentIndicator = p_Excel.DisplayCommentIndicator
	End Property

	Public Property Get DisplayDocumentActionTaskPane()
		DisplayDocumentActionTaskPane = p_Excel.DisplayDocumentActionTaskPane
	End Property

	Public Property Get DisplayDocumentInformationPanel()
		DisplayDocumentInformationPanel = p_Excel.DisplayDocumentInformationPanel
	End Property

	Public Property Get DisplayExcel4Menus()
		DisplayExcel4Menus = p_Excel.DisplayExcel4Menus
	End Property

	Public Property Get DisplayFormulaAutoComplete()
		DisplayFormulaAutoComplete = p_Excel.DisplayFormulaAutoComplete
	End Property

	Public Property Get DisplayFormulaBar()
		DisplayFormulaBar = p_Excel.DisplayFormulaBar
	End Property

	Public Property Get DisplayFullScreen()
		DisplayFullScreen = p_Excel.DisplayFullScreen
	End Property

	Public Property Get DisplayFunctionToolTips()
		DisplayFunctionToolTips = p_Excel.DisplayFunctionToolTips
	End Property

	Public Property Get DisplayInsertOptions()
		DisplayInsertOptions = p_Excel.DisplayInsertOptions
	End Property

	Public Property Get DisplayNoteIndicator()
		DisplayNoteIndicator = p_Excel.DisplayNoteIndicator
	End Property

	Public Property Get DisplayPasteOptions()
		DisplayPasteOptions = p_Excel.DisplayPasteOptions
	End Property

	Public Property Get DisplayRecentFiles()
		DisplayRecentFiles = p_Excel.DisplayRecentFiles
	End Property

	Public Property Get DisplayScrollBars()
		DisplayScrollBars = p_Excel.DisplayScrollBars
	End Property

	Public Property Get DisplayStatusBar()
		DisplayStatusBar = p_Excel.DisplayStatusBar
	End Property

	Public Property Get EditDirectlyInCell()
		EditDirectlyInCell = p_Excel.EditDirectlyInCell
	End Property

	Public Property Get EnableAutoComplete()
		EnableAutoComplete = p_Excel.EnableAutoComplete
	End Property

	Public Property Get EnableCancelKey()
		EnableCancelKey = p_Excel.EnableCancelKey
	End Property

	Public Property Get EnableCheckFileExtensions()
		EnableCheckFileExtensions = p_Excel.EnableCheckFileExtensions
	End Property

	Public Property Get EnableEvents()
		EnableEvents = p_Excel.EnableEvents
	End Property

	Public Property Get EnableLargeOperationAlert()
		EnableLargeOperationAlert = p_Excel.EnableLargeOperationAlert
	End Property

	Public Property Get EnableLivePreview()
		EnableLivePreview = p_Excel.EnableLivePreview
	End Property

	Public Property Get EnableMacroAnimations()
		EnableMacroAnimations = p_Excel.EnableMacroAnimations
	End Property

	Public Property Get EnableSound()
		EnableSound = p_Excel.EnableSound
	End Property

	Public Property Get ErrorCheckingOptions()
		ErrorCheckingOptions = p_Excel.ErrorCheckingOptions
	End Property

	Public Property Get Excel4IntlMacroSheets()
		Excel4IntlMacroSheets = p_Excel.Excel4IntlMacroSheets
	End Property

	Public Property Get Excel4MacroSheets()
		Excel4MacroSheets = p_Excel.Excel4MacroSheets
	End Property

	Public Property Get ExtendList()
		ExtendList = p_Excel.ExtendList
	End Property

	Public Property Get FeatureInstall()
		FeatureInstall = p_Excel.FeatureInstall
	End Property

	Public Property Get FileConverters()
		FileConverters = p_Excel.FileConverters
	End Property

	Public Property Get FileDialog()
		FileDialog = p_Excel.FileDialog
	End Property

	Public Property Get FileExportConverters()
		FileExportConverters = p_Excel.FileExportConverters
	End Property

	Public Property Get FileValidation()
		FileValidation = p_Excel.FileValidation
	End Property

	Public Property Get FileValidationPivot()
		FileValidationPivot = p_Excel.FileValidationPivot
	End Property

	Public Property Get FindFormat()
		FindFormat = p_Excel.FindFormat
	End Property

	Public Property Get FixedDecimal()
		FixedDecimal = p_Excel.FixedDecimal
	End Property

	Public Property Get FixedDecimalPlaces()
		FixedDecimalPlaces = p_Excel.FixedDecimalPlaces
	End Property

	Public Property Get FlashFill()
		FlashFill = p_Excel.FlashFill
	End Property

	Public Property Get FlashFillMode()
		FlashFillMode = p_Excel.FlashFillMode
	End Property

	Public Property Get FormulaBarHeight()
		FormulaBarHeight = p_Excel.FormulaBarHeight
	End Property

	Public Property Get GenerateGetPivotData()
		GenerateGetPivotData = p_Excel.GenerateGetPivotData
	End Property

	Public Property Get GenerateTableRefs()
		GenerateTableRefs = p_Excel.GenerateTableRefs
	End Property

	Public Property Get Height()
		Height = p_Excel.Height
	End Property

	Public Property Get HighQualityModeForGraphics()
		HighQualityModeForGraphics = p_Excel.HighQualityModeForGraphics
	End Property

	Public Property Get Hinstance()
		Hinstance = p_Excel.Hinstance
	End Property

	Public Property Get HinstancePtr()
		HinstancePtr = p_Excel.HinstancePtr
	End Property

	Public Property Get Hwnd()
		Hwnd = p_Excel.Hwnd
	End Property

	Public Property Get IgnoreRemoteRequests()
		IgnoreRemoteRequests = p_Excel.IgnoreRemoteRequests
	End Property

	Public Property Get Interactive()
		Interactive = p_Excel.Interactive
	End Property

	Public Property Get International()
		Interactive = p_Excel.International
	End Property

	Public Property Get IsSandboxed()
		IsSandboxed = p_Excel.IsSandboxed
	End Property

	Public Property Get Iteration()
		Iteration = p_Excel.Iteration
	End Property

	Public Property Get LanguageSettings()
		LanguageSettings = p_Excel.LanguageSettings
	End Property

	Public Property Get LargeOperationCellThousandCount()
		LargeOperationCellThousandCount = p_Excel.LargeOperationCellThousandCount
	End Property

	Public Property Get Left()
		Left = p_Excel.Left
	End Property

	Public Property Get LibraryPath()
		LibraryPath = p_Excel.LibraryPath
	End Property

	Public Property Get MailSession()
		MailSession = p_Excel.MailSession
	End Property

	Public Property Get MailSystem()
		MailSystem = p_Excel.MailSystem
	End Property

	Public Property Get MapPaperSize()
		MapPaperSize = p_Excel.MapPaperSize
	End Property

	Public Property Get MathCoprocessorAvailable()
		MathCoprocessorAvailable = p_Excel.MathCoprocessorAvailable
	End Property

	Public Property Get MaxChange()
		MaxChange = p_Excel.MaxChange
	End Property

	Public Property Get MaxIterations()
		MaxIterations = p_Excel.MaxIterations
	End Property

	Public Property Get MeasurementUnit()
		MeasurementUnit = p_Excel.MeasurementUnit
	End Property

	Public Property Get MergeInstances()
		MergeInstances = p_Excel.MergeInstances
	End Property

	Public Property Get MouseAvailable()
		MouseAvailable = p_Excel.MouseAvailable
	End Property

	Public Property Get MoveAfterReturn()
		MoveAfterReturn = p_Excel.MoveAfterReturn
	End Property

	Public Property Get MoveAfterReturnDirection()
		MoveAfterReturnDirection = p_Excel.MoveAfterReturnDirection
	End Property

	Public Property Get MultiThreadedCalculation()
		MultiThreadedCalculation = p_Excel.MultiThreadedCalculation
	End Property

	Public Property Get Name()
		Name = p_Excel.Name
	End Property

	Public Property Get Names()
		Names = p_Excel.Names
	End Property

	Public Property Get NetworkTemplatesPath()
		NetworkTemplatesPath = p_Excel.NetworkTemplatesPath
	End Property

	Public Property Get NewWorkbook()
		NewWorkbook = p_Excel.NewWorkbook
	End Property

	Public Property Get ODBCErrors()
		ODBCErrors = p_Excel.ODBCErrors
	End Property

	Public Property Get ODBCTimeout()
		ODBCTimeout = p_Excel.ODBCTimeout
	End Property

	Public Property Get OLEDBErrors()
		OLEDBErrors = p_Excel.OLEDBErrors
	End Property

	Public Property Get OnWindow()
		OnWindow = p_Excel.OnWindow
	End Property

	Public Property Get OperatingSystem()
		OperatingSystem = p_Excel.OperatingSystem
	End Property

	Public Property Get OrganizationName()
		OrganizationName = p_Excel.OrganizationName
	End Property

	Public Property Get Parent()
		Parent = p_Excel.Parent
	End Property

	Public Property Get Path()
		Path = p_Excel.Path
	End Property

	Public Property Get PathSeparator()
		PathSeparator = p_Excel.PathSeparator
	End Property

	Public Property Get PivotTableSelection()
		PivotTableSelection = p_Excel.PivotTableSelection
	End Property

	Public Property Get PreviousSelections()
		PreviousSelections = p_Excel.PreviousSelections
	End Property

	Public Property Get PrintCommunication()
		PrintCommunication = p_Excel.PrintCommunication
	End Property

	Public Property Get ProductCode()
		ProductCode = p_Excel.ProductCode
	End Property

	Public Property Get PromptForSummaryInfo()
		PromptForSummaryInfo = p_Excel.PromptForSummaryInfo
	End Property

	Public Property Get ProtectedViewWindows()
		ProtectedViewWindows = p_Excel.ProtectedViewWindows
	End Property

	Public Property Get QuickAnalysis()
		QuickAnalysis = p_Excel.QuickAnalysis
	End Property

	Public Property Get Range()
		Range = p_Excel.Range
	End Property

	Public Property Get Ready()
		Ready = p_Excel.Ready
	End Property

	Public Property Get RecentFiles()
		RecentFiles = p_Excel.RecentFiles
	End Property

	Public Property Get RecordRelative()
		RecordRelative = p_Excel.RecordRelative
	End Property

	Public Property Get ReferenceStyle()
		ReferenceStyle = p_Excel.ReferenceStyle
	End Property

	Public Property Get RegisteredFunctions()
		RegisteredFunctions = p_Excel.RegisteredFunctions
	End Property

	Public Property Get ReplaceFormat()
		ReplaceFormat = p_Excel.ReplaceFormat
	End Property

	Public Property Get RollZoom()
		RollZoom = p_Excel.RollZoom
	End Property

	Public Property Get Rows()
		Rows = p_Excel.Rows
	End Property

	Public Property Get RTD()
		RTD = p_Excel.RTD
	End Property

	Public Property Get ScreenUpdating()
		ScreenUpdating = p_Excel.ScreenUpdating
	End Property

	Public Property Get Selection()
		Selection = p_Excel.Selection
	End Property

	Public Property Get Sheets()
		Sheets = p_Excel.Sheets
	End Property

	Public Property Get SheetsInNewWorkbook()
		SheetsInNewWorkbook = p_Excel.SheetsInNewWorkbook
	End Property

	Public Property Get ShowChartTipNames()
		ShowChartTipNames = p_Excel.ShowChartTipNames
	End Property

	Public Property Get ShowChartTipValues()
		ShowChartTipValues = p_Excel.ShowChartTipValues
	End Property

	Public Property Get ShowDevTools()
		ShowDevTools = p_Excel.ShowDevTools
	End Property

	Public Property Get ShowMenuFloaties()
		ShowMenuFloaties = p_Excel.ShowMenuFloaties
	End Property

	Public Property Get ShowQuickAnalysis()
		ShowQuickAnalysis = p_Excel.ShowQuickAnalysis
	End Property

	Public Property Get ShowSelectionFloaties()
		ShowSelectionFloaties = p_Excel.ShowSelectionFloaties
	End Property

	Public Property Get ShowStartupDialog()
		ShowStartupDialog = p_Excel.ShowStartupDialog
	End Property

	Public Property Get ShowToolTips()
		ShowToolTips = p_Excel.ShowToolTips
	End Property

	Public Property Get SmartArtColors()
		SmartArtColors = p_Excel.SmartArtColors
	End Property

	Public Property Get SmartArtLayouts()
		SmartArtLayouts = p_Excel.SmartArtLayouts
	End Property

	Public Property Get SmartArtQuickStyles()
		SmartArtQuickStyles = p_Excel.SmartArtQuickStyles
	End Property

	Public Property Get Speech()
		Speech = p_Excel.Speech
	End Property

	Public Property Get SpellingOptions()
		SpellingOptions = p_Excel.SpellingOptions
	End Property

	Public Property Get StandardFont()
		StandardFont = p_Excel.StandardFont
	End Property

	Public Property Get StandardFontSize()
		StandardFontSize = p_Excel.StandardFontSize
	End Property

	Public Property Get StartupPath()
		StartupPath = p_Excel.StartupPath
	End Property

	Public Property Get StatusBar()
		StatusBar = p_Excel.StatusBar
	End Property

	Public Property Get TemplatesPath()
		TemplatesPath = p_Excel.TemplatesPath
	End Property

	Public Property Get ThisCell()
		ThisCell = p_Excel.ThisCell
	End Property

	Public Property Get ThisWorkbook()
		ThisWorkbook = p_Excel.ThisWorkbook
	End Property

	Public Property Get ThousandsSeparator()
		ThousandsSeparator = p_Excel.ThousandsSeparator
	End Property

	Public Property Get Top()
		Top = p_Excel.Top
	End Property

	Public Property Get TransitionMenuKey()
		TransitionMenuKey = p_Excel.TransitionMenuKey
	End Property

	Public Property Get TransitionMenuKeyAction()
		TransitionMenuKeyAction = p_Excel.TransitionMenuKeyAction
	End Property

	Public Property Get TransitionNavigKeys()
		TransitionNavigKeys = p_Excel.TransitionNavigKeys
	End Property

	Public Property Get UsableHeight()
		UsableHeight = p_Excel.UsableHeight
	End Property

	Public Property Get UsableWidth()
		UsableWidth = p_Excel.UsableWidth
	End Property

	Public Property Get UseClusterConnector()
		UseClusterConnector = p_Excel.UseClusterConnector
	End Property

	Public Property Get UsedObjects()
		UsedObjects = p_Excel.UsedObjects
	End Property

	Public Property Get UserControl()
		UserControl = p_Excel.UserControl
	End Property

	Public Property Get UserLibraryPath()
		UserLibraryPath = p_Excel.UserLibraryPath
	End Property

	Public Property Get UserName()
		UserName = p_Excel.UserName
	End Property

	Public Property Get UseSystemSeparators()
		UseSystemSeparators = p_Excel.UseSystemSeparators
	End Property

	Public Property Get Value()
		Value = p_Excel.Value
	End Property

	Public Property Get VBE()
		VBE = p_Excel.VBE
	End Property

	Public Property Get Version()
		Version = p_Excel.Version
	End Property

	Public Property Get Visible()
		Visible = p_Excel.Visible
	End Property

	Public Property Get WarnOnFunctionNameConflict()
		WarnOnFunctionNameConflict = p_Excel.WarnOnFunctionNameConflict
	End Property

	Public Property Get Watches()
		Watches = p_Excel.Watches
	End Property

	Public Property Get Width()
		Width = p_Excel.Width
	End Property

	Public Property Get Windows()
		Windows = p_Excel.Windows
	End Property

	Public Property Get WindowsForPens()
		WindowsForPens = p_Excel.WindowsForPens
	End Property

	Public Property Get WindowState()
		WindowState = p_Excel.WindowState
	End Property

	Public Property Get Workbooks()
		Set Workbooks = p_Excel.Workbooks
	End Property

	Public Property Get WorksheetFunction()
		WorksheetFunction = p_Excel.WorksheetFunction
	End Property

	Public Property Get Worksheets()
		Set Worksheets = p_Excel.Worksheets
	End Property


	' Methods


	Public Sub ActivateMicrosoftApp(objIndex)

	End Sub

	Public Sub AddCustomList(arrList, intByRow)

	End Sub

	Public Sub Calculate()

	End Sub

	Public Sub CalculateFull()

	End Sub

	Public Sub CalculateFullRebuild()

	End Sub

	Public Sub CalculateUntilAsyncQueriesDone()

	End Sub

	' Returns double.
	Public Function CentimetersToPoints(dblCentimeters)

	End Function

	Public Sub CheckAbort(blnKeepAbort)

	End Sub

	' Returns boolean.
	Public Function CheckSpelling(strWord, objCustDict, blnIgnoreUpper)

	End Function

	Public Function ConvertFormula() ' (Formula, FromReferenceStyle As XlReferenceStyle, [ToReferenceStyle], [ToAbsolute], [RelativeTo])

	End Function

	Public Sub DDEExecute() ' (Channel As Long, String As String)

	End Sub

	Public Function DDEInitiate() ' (App As String, Topic As String) As Long

	End Function

	Public Sub DDEPoke() ' (Channel As Long, Item, Data)

	End Sub

	Public Function DDERequest() ' (Channel As Long, Item As String)

	End Function

	Public Sub DDETerminate() ' (Channel As Long)

	End Sub

	Public Sub DeleteCustomList() ' (ListNum As Long)

	End Sub

	Public Sub DisplayXMLSourcePane() ' ([XmlMap])

	End Sub

	Public Sub DoubleClick()

	End Sub

	Public Function Evaluate() ' (Name)

	End Function

	Public Function ExecuteExcel4Macro() ' (String As String)

	End Function

	Public Function FindFile() '()  As Boolean

	End Function

	Public Function GetCustomListContents() ' (ListNum As Long)

	End Function

	Public Function GetCustomListNum() ' (ListArray) As Long

	End Function

	Public Function GetOpenFilename() ' ([FileFilter], [FilterIndex], [Title], [ButtonText], [MultiSelect])

	End Function

	Public Function GetPhonetic() ' ([Text]) As String

	End Function

	Public Function GetSaveAsFilename() ' ([InitialFilename], [FileFilter], [FilterIndex], [Title], [ButtonText])

	End Function

	' Public Sub Goto() ' ([Reference], [Scroll])

	' End Sub

	Public Sub Help() ' ([HelpFile], [HelpContextID])

	End Sub

	Public Function InchesToPoints() ' (Inches As Double) As Double

	End Function

	Public Function InputBox() ' (Prompt As String, [Title], [Default], [Left], [Top], [HelpFile], [HelpContextID], [Type])

	End Function

	Public Function Intersect() ' (Arg1 As Range, Arg2 As Range, [Arg3], [Arg4], [Arg5], [Arg6], [Arg7], [Arg8], [Arg9], [Arg10], [Arg11], [Arg12], [Arg13], [Arg14], [Arg15], [Arg16], [Arg17], [Arg18], [Arg19], [Arg20], [Arg21], [Arg22], [Arg23], [Arg24], [Arg25], [Arg26], [Arg27], [Arg28], [Arg29], [Arg30]) As Range

	End Function

	Public Sub MacroOptions() ' ([Macro], [Description], [HasMenu], [MenuText], [HasShortcutKey], [ShortcutKey], [Category], [StatusBar], [HelpContextID], [HelpFile], [ArgumentDescriptions])

	End Sub

	Public Sub MailLogoff()

	End Sub

	Public Sub MailLogon() ' ([Name], [Password], [DownloadNewMail])

	End Sub

	Public Function NextLetter() ' () As Workbook

	End Function

	Public Sub OnKey() ' (Key As String, [Procedure])

	End Sub

	Public Sub OnRepeat() ' (Text As String, Procedure As String)

	End Sub

	Public Sub OnTime() ' (EarliestTime, Procedure As String, [LatestTime], [Schedule])

	End Sub

	Public Sub OnUndo() ' (Text As String, Procedure As String)

	End Sub

	Public Sub Quit()
		p_Excel.Quit()
	End Sub

	Public Sub RecordMacro() ' ([BasicCode], [XlmCode])

	End Sub

	Public Function RegisterXLL() ' (Filename As String) As Boolean

	End Function

	Public Sub Repeat()

	End Sub

	Public Function Run() ' ([Macro], [Arg1], [Arg2], [Arg3], [Arg4], [Arg5], [Arg6], [Arg7], [Arg8], [Arg9], [Arg10], [Arg11], [Arg12], [Arg13], [Arg14], [Arg15], [Arg16], [Arg17], [Arg18], [Arg19], [Arg20], [Arg21], [Arg22], [Arg23], [Arg24], [Arg25], [Arg26], [Arg27], [Arg28], [Arg29], [Arg30])

	End Function

	Public Sub SendKeys() ' (Keys, [Wait])

	End Sub

	Public Function SharePointVersion() ' (bstrUrl As String) As Long

	End Function

	Public Sub Undo()

	End Sub

	Public Function Union() ' (Arg1 As Range, Arg2 As Range, [Arg3], [Arg4], [Arg5], [Arg6], [Arg7], [Arg8], [Arg9], [Arg10], [Arg11], [Arg12], [Arg13], [Arg14], [Arg15], [Arg16], [Arg17], [Arg18], [Arg19], [Arg20], [Arg21], [Arg22], [Arg23], [Arg24], [Arg25], [Arg26], [Arg27], [Arg28], [Arg29], [Arg30]) As Range

	End Function

	Public Sub Volatile() ' ([Volatile])

	End Sub

	Public Function Wait() ' (Time) As Boolean

	End Function


	' Termination


	Private Sub Class_Terminate()
		Set p_Excel = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Excel.vbs" Then
	Dim excel
	Set excel = New base_Excel

	WScript.Echo TypeName(excel.Workbooks)

	Set excel = Nothing
End If
