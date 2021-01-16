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

Class v_Excel
	Private pExcel


	' Initialization


	Private Sub Class_Initialize()
		Set pExcel = CreateObject("Excel.Application")
	End Sub


	' Properties


	Public Property Get ActiveCell()
		ActiveCell = pExcel.ActiveCell
	End Property

	Public Property Get ActiveChart()
		ActiveChart = pExcel.ActiveChart
	End Property

	Public Property Get ActiveEncryptionSession()
		ActiveEncryptionSession = pExcel.ActiveEncryptionSession
	End Property

	Public Property Get ActivePrinter()
		ActivePrinter = pExcel.ActivePrinter
	End Property

	Public Property Get ActiveProtectedViewWindow()
		ActiveProtectedViewWindow = pExcel.ActiveProtectedViewWindow
	End Property

	Public Property Get ActiveSheet()
		ActiveSheet = pExcel.ActiveSheet
	End Property

	Public Property Get ActiveWindow()
		ActiveWindow = pExcel.ActiveWindow
	End Property

	Public Property Get ActiveWorkbook()
		ActiveWorkbook = pExcel.ActiveWorkbook
	End Property

	Public Property Get AddIns()
		AddIns = pExcel.AddIns
	End Property

	Public Property Get AddIns2()
		AddIns2 = pExcel.AddIns2
	End Property

	Public Property Get AlertBeforeOverwriting()
		AlertBeforeOverwriting = pExcel.AlertBeforeOverwriting
	End Property

	Public Property Get AltStartupPath()
		AltStartupPath = pExcel.AltStartupPath
	End Property

	Public Property Get AlwaysUseClearType()
		AlwaysUseClearType = pExcel.AlwaysUseClearType
	End Property

	Public Property Get Application()
		Application = pExcel.Application
	End Property

	Public Property Get ArbitraryXMLSupportAvailable()
		ArbitraryXMLSupportAvailable = pExcel.ArbitraryXMLSupportAvailable
	End Property

	Public Property Get AskToUpdateLinks()
		AskToUpdateLinks = pExcel.AskToUpdateLinks
	End Property

	Public Property Get Assistance()
		Assistance = pExcel.Assistance
	End Property

	Public Property Get AutoCorrect()
		AutoCorrect = pExcel.AutoCorrect
	End Property

	Public Property Get AutoFormatAsYouTypeReplaceHyperlinks()
		AutoFormatAsYouTypeReplaceHyperlinks = pExcel.AutoFormatAsYouTypeReplaceHyperlinks
	End Property

	Public Property Get AutomationSecurity()
		AutomationSecurity = pExcel.AutomationSecurity
	End Property

	Public Property Get AutoPercentEntry()
		AutoPercentEntry = pExcel.AutoPercentEntry
	End Property

	Public Property Get AutoRecover()
		AutoRecover = pExcel.AutoRecover
	End Property

	Public Property Get Build()
		Build = pExcel.Build
	End Property

	Public Property Get CalculateBeforeSave()
		CalculateBeforeSave = pExcel.CalculateBeforeSave
	End Property

	Public Property Get Calculation()
		Calculation = pExcel.Calculation
	End Property

	Public Property Get CalculationInterruptKey()
		CalculationInterruptKey = pExcel.CalculationInterruptKey
	End Property

	Public Property Get CalculationState()
		CalculationState = pExcel.CalculationState
	End Property

	Public Property Get CalculationVersion()
		CalculationVersion = pExcel.CalculationVersion
	End Property

	Public Property Get Caller()
		Caller = pExcel.Caller
	End Property

	Public Property Get CanPlaySounds()
		CanPlaySounds = pExcel.CanPlaySounds
	End Property

	Public Property Get CanRecordSounds()
		CanRecordSounds = pExcel.CanRecordSounds
	End Property

	Public Property Get Caption()
		Caption = pExcel.Caption
	End Property

	Public Property Get CellDragAndDrop()
		CellDragAndDrop = pExcel.CellDragAndDrop
	End Property

	Public Property Get Cells()
		Cells = pExcel.Cells
	End Property

	Public Property Get ChartDataPointTrack()
		ChartDataPointTrack = pExcel.ChartDataPointTrack
	End Property

	Public Property Get Charts()
		Charts = pExcel.Charts
	End Property

	Public Property Get ClipboardFormats()
		ClipboardFormats = pExcel.ClipboardFormats
	End Property

	Public Property Get ClusterConnector()
		ClusterConnector = pExcel.ClusterConnector
	End Property

	Public Property Get Columns()
		Columns = pExcel.Columns
	End Property

	Public Property Get COMAddIns()
		COMAddIns = pExcel.COMAddIns
	End Property

	Public Property Get CommandBars()
		CommandBars = pExcel.CommandBars
	End Property

	Public Property Get CommandUnderlines()
		CommandUnderlines = pExcel.CommandUnderlines
	End Property

	Public Property Get ConstrainNumeric()
		ConstrainNumeric = pExcel.ConstrainNumeric
	End Property

	Public Property Get ControlCharacters()
		ControlCharacters = pExcel.ControlCharacters
	End Property

	Public Property Get CopyObjectsWithCells()
		CopyObjectsWithCells = pExcel.CopyObjectsWithCells
	End Property

	Public Property Get Creator()
		Creator = pExcel.Creator
	End Property

	Public Property Get Cursor()
		Cursor = pExcel.Cursor
	End Property

	Public Property Get CursorMovement()
		CursorMovement = pExcel.CursorMovement
	End Property

	Public Property Get CustomListCount()
		CustomListCount = pExcel.CustomListCount
	End Property

	Public Property Get CutCopyMode()
		CutCopyMode = pExcel.CutCopyMode
	End Property

	Public Property Get DataEntryMode()
		DataEntryMode = pExcel.DataEntryMode
	End Property

	Public Property Get DDEAppReturnCode()
		DDEAppReturnCode = pExcel.DDEAppReturnCode
	End Property

	Public Property Get DecimalSeparator()
		DecimalSeparator = pExcel.DecimalSeparator
	End Property

	Public Property Get DefaultFilePath()
		DefaultFilePath = pExcel.DefaultFilePath
	End Property

	Public Property Get DefaultSaveFormat()
		DefaultSaveFormat = pExcel.DefaultSaveFormat
	End Property

	Public Property Get DefaultSheetDirection()
		DefaultSheetDirection = pExcel.DefaultSheetDirection
	End Property

	Public Property Get DefaultWebOptions()
		DefaultWebOptions = pExcel.DefaultWebOptions
	End Property

	Public Property Get DeferAsyncQueries()
		DeferAsyncQueries = pExcel.DeferAsyncQueries
	End Property

	Public Property Get Dialogs()
		Dialogs = pExcel.Dialogs
	End Property

	Public Property Get DisplayAlerts()
		DisplayAlerts = pExcel.DisplayAlerts
	End Property

	Public Property Get DisplayClipboardWindow()
		DisplayClipboardWindow = pExcel.DisplayClipboardWindow
	End Property

	Public Property Get DisplayCommentIndicator()
		DisplayCommentIndicator = pExcel.DisplayCommentIndicator
	End Property

	Public Property Get DisplayDocumentActionTaskPane()
		DisplayDocumentActionTaskPane = pExcel.DisplayDocumentActionTaskPane
	End Property

	Public Property Get DisplayDocumentInformationPanel()
		DisplayDocumentInformationPanel = pExcel.DisplayDocumentInformationPanel
	End Property

	Public Property Get DisplayExcel4Menus()
		DisplayExcel4Menus = pExcel.DisplayExcel4Menus
	End Property

	Public Property Get DisplayFormulaAutoComplete()
		DisplayFormulaAutoComplete = pExcel.DisplayFormulaAutoComplete
	End Property

	Public Property Get DisplayFormulaBar()
		DisplayFormulaBar = pExcel.DisplayFormulaBar
	End Property

	Public Property Get DisplayFullScreen()
		DisplayFullScreen = pExcel.DisplayFullScreen
	End Property

	Public Property Get DisplayFunctionToolTips()
		DisplayFunctionToolTips = pExcel.DisplayFunctionToolTips
	End Property

	Public Property Get DisplayInsertOptions()
		DisplayInsertOptions = pExcel.DisplayInsertOptions
	End Property

	Public Property Get DisplayNoteIndicator()
		DisplayNoteIndicator = pExcel.DisplayNoteIndicator
	End Property

	Public Property Get DisplayPasteOptions()
		DisplayPasteOptions = pExcel.DisplayPasteOptions
	End Property

	Public Property Get DisplayRecentFiles()
		DisplayRecentFiles = pExcel.DisplayRecentFiles
	End Property

	Public Property Get DisplayScrollBars()
		DisplayScrollBars = pExcel.DisplayScrollBars
	End Property

	Public Property Get DisplayStatusBar()
		DisplayStatusBar = pExcel.DisplayStatusBar
	End Property

	Public Property Get EditDirectlyInCell()
		EditDirectlyInCell = pExcel.EditDirectlyInCell
	End Property

	Public Property Get EnableAutoComplete()
		EnableAutoComplete = pExcel.EnableAutoComplete
	End Property

	Public Property Get EnableCancelKey()
		EnableCancelKey = pExcel.EnableCancelKey
	End Property

	Public Property Get EnableCheckFileExtensions()
		EnableCheckFileExtensions = pExcel.EnableCheckFileExtensions
	End Property

	Public Property Get EnableEvents()
		EnableEvents = pExcel.EnableEvents
	End Property

	Public Property Get EnableLargeOperationAlert()
		EnableLargeOperationAlert = pExcel.EnableLargeOperationAlert
	End Property

	Public Property Get EnableLivePreview()
		EnableLivePreview = pExcel.EnableLivePreview
	End Property

	Public Property Get EnableMacroAnimations()
		EnableMacroAnimations = pExcel.EnableMacroAnimations
	End Property

	Public Property Get EnableSound()
		EnableSound = pExcel.EnableSound
	End Property

	Public Property Get ErrorCheckingOptions()
		ErrorCheckingOptions = pExcel.ErrorCheckingOptions
	End Property

	Public Property Get Excel4IntlMacroSheets()
		Excel4IntlMacroSheets = pExcel.Excel4IntlMacroSheets
	End Property

	Public Property Get Excel4MacroSheets()
		Excel4MacroSheets = pExcel.Excel4MacroSheets
	End Property

	Public Property Get ExtendList()
		ExtendList = pExcel.ExtendList
	End Property

	Public Property Get FeatureInstall()
		FeatureInstall = pExcel.FeatureInstall
	End Property

	Public Property Get FileConverters()
		FileConverters = pExcel.FileConverters
	End Property

	Public Property Get FileDialog()
		FileDialog = pExcel.FileDialog
	End Property

	Public Property Get FileExportConverters()
		FileExportConverters = pExcel.FileExportConverters
	End Property

	Public Property Get FileValidation()
		FileValidation = pExcel.FileValidation
	End Property

	Public Property Get FileValidationPivot()
		FileValidationPivot = pExcel.FileValidationPivot
	End Property

	Public Property Get FindFormat()
		FindFormat = pExcel.FindFormat
	End Property

	Public Property Get FixedDecimal()
		FixedDecimal = pExcel.FixedDecimal
	End Property

	Public Property Get FixedDecimalPlaces()
		FixedDecimalPlaces = pExcel.FixedDecimalPlaces
	End Property

	Public Property Get FlashFill()
		FlashFill = pExcel.FlashFill
	End Property

	Public Property Get FlashFillMode()
		FlashFillMode = pExcel.FlashFillMode
	End Property

	Public Property Get FormulaBarHeight()
		FormulaBarHeight = pExcel.FormulaBarHeight
	End Property

	Public Property Get GenerateGetPivotData()
		GenerateGetPivotData = pExcel.GenerateGetPivotData
	End Property

	Public Property Get GenerateTableRefs()
		GenerateTableRefs = pExcel.GenerateTableRefs
	End Property

	Public Property Get Height()
		Height = pExcel.Height
	End Property

	Public Property Get HighQualityModeForGraphics()
		HighQualityModeForGraphics = pExcel.HighQualityModeForGraphics
	End Property

	Public Property Get Hinstance()
		Hinstance = pExcel.Hinstance
	End Property

	Public Property Get HinstancePtr()
		HinstancePtr = pExcel.HinstancePtr
	End Property

	Public Property Get Hwnd()
		Hwnd = pExcel.Hwnd
	End Property

	Public Property Get IgnoreRemoteRequests()
		IgnoreRemoteRequests = pExcel.IgnoreRemoteRequests
	End Property

	Public Property Get Interactive()
		Interactive = pExcel.Interactive
	End Property

	Public Property Get International()
		Interactive = pExcel.International
	End Property

	Public Property Get IsSandboxed()
		IsSandboxed = pExcel.IsSandboxed
	End Property

	Public Property Get Iteration()
		Iteration = pExcel.Iteration
	End Property

	Public Property Get LanguageSettings()
		LanguageSettings = pExcel.LanguageSettings
	End Property

	Public Property Get LargeOperationCellThousandCount()
		LargeOperationCellThousandCount = pExcel.LargeOperationCellThousandCount
	End Property

	Public Property Get Left()
		Left = pExcel.Left
	End Property

	Public Property Get LibraryPath()
		LibraryPath = pExcel.LibraryPath
	End Property

	Public Property Get MailSession()
		MailSession = pExcel.MailSession
	End Property

	Public Property Get MailSystem()
		MailSystem = pExcel.MailSystem
	End Property

	Public Property Get MapPaperSize()
		MapPaperSize = pExcel.MapPaperSize
	End Property

	Public Property Get MathCoprocessorAvailable()
		MathCoprocessorAvailable = pExcel.MathCoprocessorAvailable
	End Property

	Public Property Get MaxChange()
		MaxChange = pExcel.MaxChange
	End Property

	Public Property Get MaxIterations()
		MaxIterations = pExcel.MaxIterations
	End Property

	Public Property Get MeasurementUnit()
		MeasurementUnit = pExcel.MeasurementUnit
	End Property

	Public Property Get MergeInstances()
		MergeInstances = pExcel.MergeInstances
	End Property

	Public Property Get MouseAvailable()
		MouseAvailable = pExcel.MouseAvailable
	End Property

	Public Property Get MoveAfterReturn()
		MoveAfterReturn = pExcel.MoveAfterReturn
	End Property

	Public Property Get MoveAfterReturnDirection()
		MoveAfterReturnDirection = pExcel.MoveAfterReturnDirection
	End Property

	Public Property Get MultiThreadedCalculation()
		MultiThreadedCalculation = pExcel.MultiThreadedCalculation
	End Property

	Public Property Get Name()
		Name = pExcel.Name
	End Property

	Public Property Get Names()
		Names = pExcel.Names
	End Property

	Public Property Get NetworkTemplatesPath()
		NetworkTemplatesPath = pExcel.NetworkTemplatesPath
	End Property

	Public Property Get NewWorkbook()
		NewWorkbook = pExcel.NewWorkbook
	End Property

	Public Property Get ODBCErrors()
		ODBCErrors = pExcel.ODBCErrors
	End Property

	Public Property Get ODBCTimeout()
		ODBCTimeout = pExcel.ODBCTimeout
	End Property

	Public Property Get OLEDBErrors()
		OLEDBErrors = pExcel.OLEDBErrors
	End Property

	Public Property Get OnWindow()
		OnWindow = pExcel.OnWindow
	End Property

	Public Property Get OperatingSystem()
		OperatingSystem = pExcel.OperatingSystem
	End Property

	Public Property Get OrganizationName()
		OrganizationName = pExcel.OrganizationName
	End Property

	Public Property Get Parent()
		Parent = pExcel.Parent
	End Property

	Public Property Get Path()
		Path = pExcel.Path
	End Property

	Public Property Get PathSeparator()
		PathSeparator = pExcel.PathSeparator
	End Property

	Public Property Get PivotTableSelection()
		PivotTableSelection = pExcel.PivotTableSelection
	End Property

	Public Property Get PreviousSelections()
		PreviousSelections = pExcel.PreviousSelections
	End Property

	Public Property Get PrintCommunication()
		PrintCommunication = pExcel.PrintCommunication
	End Property

	Public Property Get ProductCode()
		ProductCode = pExcel.ProductCode
	End Property

	Public Property Get PromptForSummaryInfo()
		PromptForSummaryInfo = pExcel.PromptForSummaryInfo
	End Property

	Public Property Get ProtectedViewWindows()
		ProtectedViewWindows = pExcel.ProtectedViewWindows
	End Property

	Public Property Get QuickAnalysis()
		QuickAnalysis = pExcel.QuickAnalysis
	End Property

	Public Property Get Range()
		Range = pExcel.Range
	End Property

	Public Property Get Ready()
		Ready = pExcel.Ready
	End Property

	Public Property Get RecentFiles()
		RecentFiles = pExcel.RecentFiles
	End Property

	Public Property Get RecordRelative()
		RecordRelative = pExcel.RecordRelative
	End Property

	Public Property Get ReferenceStyle()
		ReferenceStyle = pExcel.ReferenceStyle
	End Property

	Public Property Get RegisteredFunctions()
		RegisteredFunctions = pExcel.RegisteredFunctions
	End Property

	Public Property Get ReplaceFormat()
		ReplaceFormat = pExcel.ReplaceFormat
	End Property

	Public Property Get RollZoom()
		RollZoom = pExcel.RollZoom
	End Property

	Public Property Get Rows()
		Rows = pExcel.Rows
	End Property

	Public Property Get RTD()
		RTD = pExcel.RTD
	End Property

	Public Property Get ScreenUpdating()
		ScreenUpdating = pExcel.ScreenUpdating
	End Property

	Public Property Get Selection()
		Selection = pExcel.Selection
	End Property

	Public Property Get Sheets()
		Sheets = pExcel.Sheets
	End Property

	Public Property Get SheetsInNewWorkbook()
		SheetsInNewWorkbook = pExcel.SheetsInNewWorkbook
	End Property

	Public Property Get ShowChartTipNames()
		ShowChartTipNames = pExcel.ShowChartTipNames
	End Property

	Public Property Get ShowChartTipValues()
		ShowChartTipValues = pExcel.ShowChartTipValues
	End Property

	Public Property Get ShowDevTools()
		ShowDevTools = pExcel.ShowDevTools
	End Property

	Public Property Get ShowMenuFloaties()
		ShowMenuFloaties = pExcel.ShowMenuFloaties
	End Property

	Public Property Get ShowQuickAnalysis()
		ShowQuickAnalysis = pExcel.ShowQuickAnalysis
	End Property

	Public Property Get ShowSelectionFloaties()
		ShowSelectionFloaties = pExcel.ShowSelectionFloaties
	End Property

	Public Property Get ShowStartupDialog()
		ShowStartupDialog = pExcel.ShowStartupDialog
	End Property

	Public Property Get ShowToolTips()
		ShowToolTips = pExcel.ShowToolTips
	End Property

	Public Property Get SmartArtColors()
		SmartArtColors = pExcel.SmartArtColors
	End Property

	Public Property Get SmartArtLayouts()
		SmartArtLayouts = pExcel.SmartArtLayouts
	End Property

	Public Property Get SmartArtQuickStyles()
		SmartArtQuickStyles = pExcel.SmartArtQuickStyles
	End Property

	Public Property Get Speech()
		Speech = pExcel.Speech
	End Property

	Public Property Get SpellingOptions()
		SpellingOptions = pExcel.SpellingOptions
	End Property

	Public Property Get StandardFont()
		StandardFont = pExcel.StandardFont
	End Property

	Public Property Get StandardFontSize()
		StandardFontSize = pExcel.StandardFontSize
	End Property

	Public Property Get StartupPath()
		StartupPath = pExcel.StartupPath
	End Property

	Public Property Get StatusBar()
		StatusBar = pExcel.StatusBar
	End Property

	Public Property Get TemplatesPath()
		TemplatesPath = pExcel.TemplatesPath
	End Property

	Public Property Get ThisCell()
		ThisCell = pExcel.ThisCell
	End Property

	Public Property Get ThisWorkbook()
		ThisWorkbook = pExcel.ThisWorkbook
	End Property

	Public Property Get ThousandsSeparator()
		ThousandsSeparator = pExcel.ThousandsSeparator
	End Property

	Public Property Get Top()
		Top = pExcel.Top
	End Property

	Public Property Get TransitionMenuKey()
		TransitionMenuKey = pExcel.TransitionMenuKey
	End Property

	Public Property Get TransitionMenuKeyAction()
		TransitionMenuKeyAction = pExcel.TransitionMenuKeyAction
	End Property

	Public Property Get TransitionNavigKeys()
		TransitionNavigKeys = pExcel.TransitionNavigKeys
	End Property

	Public Property Get UsableHeight()
		UsableHeight = pExcel.UsableHeight
	End Property

	Public Property Get UsableWidth()
		UsableWidth = pExcel.UsableWidth
	End Property

	Public Property Get UseClusterConnector()
		UseClusterConnector = pExcel.UseClusterConnector
	End Property

	Public Property Get UsedObjects()
		UsedObjects = pExcel.UsedObjects
	End Property

	Public Property Get UserControl()
		UserControl = pExcel.UserControl
	End Property

	Public Property Get UserLibraryPath()
		UserLibraryPath = pExcel.UserLibraryPath
	End Property

	Public Property Get UserName()
		UserName = pExcel.UserName
	End Property

	Public Property Get UseSystemSeparators()
		UseSystemSeparators = pExcel.UseSystemSeparators
	End Property

	Public Property Get Value()
		Value = pExcel.Value
	End Property

	Public Property Get VBE()
		VBE = pExcel.VBE
	End Property

	Public Property Get Version()
		Version = pExcel.Version
	End Property

	Public Property Get Visible()
		Visible = pExcel.Visible
	End Property

	Public Property Get WarnOnFunctionNameConflict()
		WarnOnFunctionNameConflict = pExcel.WarnOnFunctionNameConflict
	End Property

	Public Property Get Watches()
		Watches = pExcel.Watches
	End Property

	Public Property Get Width()
		Width = pExcel.Width
	End Property

	Public Property Get Windows()
		Windows = pExcel.Windows
	End Property

	Public Property Get WindowsForPens()
		WindowsForPens = pExcel.WindowsForPens
	End Property

	Public Property Get WindowState()
		WindowState = pExcel.WindowState
	End Property

	Public Property Get Workbooks()
		Set Workbooks = pExcel.Workbooks
	End Property

	Public Property Get WorksheetFunction()
		WorksheetFunction = pExcel.WorksheetFunction
	End Property

	Public Property Get Worksheets()
		Set Worksheets = pExcel.Worksheets
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
		pExcel.Quit()
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
		Set pExcel = Nothing
	End Sub
End Class

If WScript.ScriptName = "v_Excel.vbs" Then
	Dim excel
	Set excel = New v_Excel

	WScript.Echo TypeName(excel.Workbooks)

	Set excel = Nothing
End If