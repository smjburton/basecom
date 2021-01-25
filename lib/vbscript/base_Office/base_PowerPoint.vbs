Option Explicit

Class base_PowerPoint
	Private p_PowerPoint

	Private Sub Class_Initialize()
		Set p_PowerPoint = CreateObject("PowerPoint.Application")
	End Sub


	' Properties


	Public Property Get Active()
		Set Active = p_PowerPoint.Active 
	End Property

	Public Property Get ActiveEncryptionSession()
		ActiveEncryptionSession = p_PowerPoint.ActiveEncryptionSession 
	End Property

	Public Property Get ActivePresentation()
		Set ActivePresentation = p_PowerPoint.ActivePresentation 
	End Property

	Public Property Get ActivePrinter()
		ActivePrinter = p_PowerPoint.ActivePrinter 
	End Property

	Public Property Get ActiveProtectedViewWindow()
		Set ActiveProtectedViewWindow = p_PowerPoint.ActiveProtectedViewWindow 
	End Property

	Public Property Get ActiveWindow()
		Set ActiveWindow = p_PowerPoint.ActiveWindow 
	End Property

	Public Property Get AddIns()
		Set AddIns = p_PowerPoint.AddIns 
	End Property

	Public Property Get Assistance()
		Set Assistance = p_PowerPoint.Assistance 
	End Property

	Public Property Get AutoCorrect()
		Set AutoCorrect = p_PowerPoint.AutoCorrect 
	End Property

	Public Property Get AutomationSecurity()
		Set AutomationSecurity = p_PowerPoint.AutomationSecurity 
	End Property

	Public Property Set AutomationSecurity(objMsoAutomationSecurity)
		Set p_PowerPoint.AutomationSecurity = objMsoAutomationSecurity
	End Property

	Public Property Get Build()
		Build = p_PowerPoint.Build 
	End Property

	Public Property Get Caption()
		Caption = p_PowerPoint.Caption 
	End Property

	Public Property Let Caption(strCaption)
		p_PowerPoint.Caption = strCaption
	End Property

	Public Property Get ChartDataPointTrack()
		ChartDataPointTrack = p_PowerPoint.ChartDataPointTrack 
	End Property

	Public Property Let ChartDataPointTrack(blnChartDataPointTrack)
		p_PowerPoint.ChartDataPointTrack = blnChartDataPointTrack
	End Property

	Public Property Get COMAddIns()
		Set COMAddIns = p_PowerPoint.COMAddIns 
	End Property

	Public Property Get CommandBars()
		Set CommandBars = p_PowerPoint.CommandBars 
	End Property

	Public Property Get Creator()
		Creator = p_PowerPoint.Creator 
	End Property

	Public Property Get DisplayAlerts()
		Set DisplayAlerts = p_PowerPoint.DisplayAlerts 
	End Property

	Public Property Set DisplayAlerts(objPpAlertLevel)
		Set p_PowerPoint.DisplayAlerts = objPpAlertLevel
	End Property

	Public Property Get DisplayDocumentInformationPanel()
		DisplayDocumentInformationPanel = p_PowerPoint.DisplayDocumentInformationPanel 
	End Property

	Public Property Let DisplayDocumentInformationPanel(blnDisplayDocumentInformationPanel)
		p_PowerPoint.DisplayDocumentInformationPanel = blnDisplayDocumentInformationPanel
	End Property

	Public Property Get DisplayGridLines()
		Set DisplayGridLines = p_PowerPoint.DisplayGridLines 
	End Property

	Public Property Set DisplayGridLines(objDisplayGridLines)
		Set p_PowerPoint.DisplayGridLines = objDisplayGridLines
	End Property

	Public Property Get DisplayGuides()
		Set DisplayGuides = p_PowerPoint.DisplayGuides 
	End Property

	Public Property Set DisplayGuides(objMsoTriState)
		Set p_PowerPoint.DisplayGuides = objMsoTriState
	End Property

	Public Property Get FeatureInstall()
		Set FeatureInstall = p_PowerPoint.FeatureInstall 
	End Property

	Public Property Set FeatureInstall(objMsoFeatureInstall)
		Set p_PowerPoint.FeatureInstall = objMsoFeatureInstall
	End Property

	Public Property Get FileConverters()
		Set FileConverters = p_PowerPoint.FileConverters 
	End Property

	Public Property Get FileDialog(objType)
		Set FileDialog = p_PowerPoint.FileDialog(objType)
	End Property

	Public Property Get FileValidation()
		Set FileValidation = p_PowerPoint.FileValidation 
	End Property

	Public Property Set FileValidation(objMsoFileValidationMode)
		Set p_PowerPoint.FileValidation = objMsoFileValidationMode
	End Property

	Public Property Get Height()
		Height = p_PowerPoint.Height 
	End Property

	Public Property Let Height(sngHeight)
		p_PowerPoint.Height = sngHeight
	End Property

	Public Property Get IsSandboxed()
		IsSandboxed = p_PowerPoint.IsSandboxed 
	End Property

	Public Property Get LanguageSettings()
		Set LanguageSettings = p_PowerPoint.LanguageSettings 
	End Property

	Public Property Get Left()
		Left = p_PowerPoint.Left 
	End Property

	Public Property Let Left(sngLeft)
		p_PowerPoint.Left = sngLeft
	End Property

	Public Default Property Get Name()
		Name = p_PowerPoint.Name 
	End Property

	Public Property Get NewPresentation()
		Set NewPresentation = p_PowerPoint.NewPresentation 
	End Property

	Public Property Get OperatingSystem()
		OperatingSystem = p_PowerPoint.OperatingSystem 
	End Property

	Public Property Get Options()
		Set Options = p_PowerPoint.Options 
	End Property

	Public Property Get Path()
		Path = p_PowerPoint.Path 
	End Property

	Public Property Get Presentations()
		Set Presentations = p_PowerPoint.Presentations 
	End Property

	Public Property Get ProductCode()
		ProductCode = p_PowerPoint.ProductCode 
	End Property

	Public Property Get ProtectedViewWindows()
		Set ProtectedViewWindows = p_PowerPoint.ProtectedViewWindows 
	End Property

	Public Property Get ShowStartupDialog()
		Set ShowStartupDialog = p_PowerPoint.ShowStartupDialog 
	End Property

	Public Property Set ShowStartupDialog(objMsoTriState)
		Set p_PowerPoint.ShowStartupDialog = objMsoTriState
	End Property

	Public Property Get ShowWindowsInTaskbar()
		Set ShowWindowsInTaskbar = p_PowerPoint.ShowWindowsInTaskbar 
	End Property

	Public Property Set ShowWindowsInTaskbar(objMsoTriState)
		Set p_PowerPoint.ShowWindowsInTaskbar = objMsoTriState
	End Property

	Public Property Get SlideShowWindows()
		Set SlideShowWindows = p_PowerPoint.SlideShowWindows 
	End Property

	Public Property Get SmartArtColors()
		Set SmartArtColors = p_PowerPoint.SmartArtColors 
	End Property

	Public Property Get SmartArtLayouts()
		Set SmartArtLayouts = p_PowerPoint.SmartArtLayouts 
	End Property

	Public Property Get SmartArtQuickStyles()
		Set SmartArtQuickStyles = p_PowerPoint.SmartArtQuickStyles 
	End Property

	Public Property Get Top()
		p_PowerPoint.Top 
	End Property

	Public Property Let Top(sngTop)
		p_PowerPoint.Top = sngTop
	End Property

	Public Property Get VBE()
		Set VBE = p_PowerPoint.VBE 
	End Property

	Public Property Get Version()
		Version = p_PowerPoint.Version 
	End Property

	Public Property Get Visible()
		Set Visible = p_PowerPoint.Visible 
	End Property

	Public Property Set Visible(objMsoTriState)
		Set p_PowerPoint.Visible = objMsoTriState
	End Property

	Public Property Get Width()
		Width = p_PowerPoint.Width 
	End Property

	Public Property Let Width(sngWidth)
		p_PowerPoint.Width = sngWidth
	End Property

	Public Property Get Windows()
		Set Windows = p_PowerPoint.Windows 
	End Property

	Public Property Get WindowState()
		Set WindowState = p_PowerPoint.WindowState 
	End Property

	Public Property Set WindowState(objPpWindowState)
		Set p_PowerPoint.WindowState = objPpWindowState
	End Property


	' Methods


	Public Sub Activate()
		p_PowerPoint.Activate
	End Sub

	Public Sub Help() ' Optional params: [HelpFile As String = "vbapp10.chm"], [ContextID As Long])
		p_PowerPoint.Help
	End Sub

	Public Function OpenThemeFile(strThemeFileName)
		Set OpenThemeFile = p_PowerPoint.OpenThemeFile(strThemeFileName)
	End Function

	Public Sub Quit()
		p_PowerPoint.Quit
	End Sub

	Public Function Run(strMacroName, varParamArray) ' safeArrayOfParams() As Variant)
		Run = p_PowerPoint.Run(strMacroName, varParamArray)
	End Function

	Public Sub StartNewUndoEntry()
		p_PowerPoint.StartNewUndoEntry
	End Sub


	' Events

	' AfterDragDropOnSlide
	' AfterNewPresentation
	' AfterPresentationOpen
	' AfterShapeSizeChange
	' ColorSchemeChanged
	' NewPresentation
	' PresentationBeforeClose
	' PresentationBeforeSave
	' PresentationClose
	' PresentationCloseFinal
	' PresentationNewSlide
	' PresentationOpen
	' PresentationPrint
	' PresentationSave
	' PresentationSync
	' ProtectedViewWindowActivate
	' ProtectedViewWindowBeforeClose
	' ProtectedViewWindowBeforeEdit
	' ProtectedViewWindowDeactivate
	' ProtectedViewWindowOpen
	' SlideSelectionChanged
	' SlideShowBegin
	' SlideShowEnd
	' SlideShowNextBuild
	' SlideShowNextClick
	' SlideShowNextSlide
	' SlideShowOnNext
	' SlideShowOnPrevious
	' WindowActivate
	' WindowBeforeDoubleClick
	' WindowBeforeRightClick
	' WindowDeactivate
	' WindowSelectionChange


	Private Sub Class_Terminate()
		Set p_PowerPoint = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_PowerPoint.vbs" Then

End If