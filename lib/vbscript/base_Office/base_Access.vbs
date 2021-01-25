Option Explicit

Class base_Access_Application
    Private p_Application

    Private Sub Class_Initialize()
        Set p_Application = CreateObject("Access.Application")
    End Sub


    ' Properties


    Public Property Get Application()
        Set Application = p_Application.Application
    End Property

    Public Property Get Assistance()
        Set Assistance = p_Application.Assistance
    End Property

    Public Property Get AutoCorrect()
        Set AutoCorrect = p_Application.AutoCorrect
    End Property

    Public Property Get AutomationSecurity()
        Set AutomationSecurity = p_Application.AutomationSecurity
    End Property

    Public Property Get BrokenReference()
        BrokenReference = p_Application.BrokenReference
    End Property

    Public Property Get Build()
        Build = p_Application.Build
    End Property

    Public Property Get CodeContextObject()
        Set CodeContextObject = p_Application.CodeContextObject
    End Property

    Public Property Get CodeData()
        Set CodeData = p_Application.CodeData
    End Property

    Public Property Get CodeProject()
        Set CodeProject = p_Application.CodeProject
    End Property

    Public Property Get COMAddIns()
        Set COMAddIns = p_Application.COMAddIns
    End Property

    Public Property Get CommandBars()
        Set CommandBars = p_Application.CommandBars
    End Property

    Public Property Get CurrentData()
        Set CurrentData = p_Application.CurrentData
    End Property

    Public Property Get CurrentObjectName()
        CurrentObjectName = p_Application.CurrentObjectName
    End Property

    Public Property Get CurrentObjectType()
        Set CurrentObjectType = p_Application.CurrentObjectType
    End Property

    Public Property Get CurrentProject()
        Set CurrentProject = p_Application.CurrentProject
    End Property

    Public Property Get DBEngine()
        Set DBEngine = p_Application.DBEngine
    End Property

    Public Property Get DoCmd()
        Set DoCmd = p_Application.DoCmd
    End Property

    Public Property Get FeatureInstall()
        Set FeatureInstall = p_Application.FeatureInstall
    End Property

    Public Property Set FeatureInstall(objMsoFeatureInstall)
        Set p_Application.FeatureInstall = objMsoFeatureInstall
    End Property

    Public Property Get FileDialog()
        Set FileDialog = p_Application.FileDialog
    End Property

    Public Property Get Forms()
        Set Forms = p_Application.Forms
    End Property

    Public Property Get IsCompiled()
        IsCompiled = p_Application.IsCompiled
    End Property

    Public Property Get LanguageSettings()
        Set LanguageSettings = p_Application.LanguageSettings
    End Property

    Public Property Get MacroError()
        Set MacroError = p_Application.MacroError
    End Property

    Public Property Get MenuBar()
        MenuBar = p_Application.MenuBar
    End Property

    Public Property Let MenuBar(strMenuBar)
        p_Application.MenuBar = strMenuBar
    End Property

    Public Property Get Modules()
        Set Modules = p_Application.Modules
    End Property

    Public Property Get Name()
        Name = p_Application.Name
    End Property

    Public Property Get NewFileTaskPane()
        Set NewFileTaskPane = p_Application.NewFileTaskPane
    End Property

    Public Property Get Parent()
        Set Parent = p_Application.Parent
    End Property

    Public Property Get Printer()
        Set Printer = p_Application.Printer
    End Property

    Public Property Set Printer(objPrinter)
        Set p_Application.Printer = objPrinter
    End Property

    Public Property Get Printers()
        Set Printers = p_Application.Printers
    End Property

    Public Property Get ProductCode()
        ProductCode = p_Application.ProductCode
    End Property

    Public Property Get References()
        Set References = p_Application.References
    End Property

    Public Property Get Reports()
        Set Reports = p_Application.Reports
    End Property

    Public Property Get ReturnVars()
        Set ReturnVars = p_Application.ReturnVars
    End Property

    Public Property Get Screen()
        Set Screen = p_Application.Screen
    End Property

    Public Property Get ShortcutMenuBar()
        ShortcutMenuBar = p_Application.ShortcutMenuBar
    End Property

    Public Property Let ShortcutMenuBar(strShortcutMenubar)
        p_Application.ShortcutMenuBar = strShortcutMenubar
    End Property

    Public Property Get TempVars()
        Set TempVars = p_Application.TempVars
    End Property

    Public Property Get UserControl()
        UserControl = p_Application.UserControl
    End Property

    Public Property Let UserControl(blnUserControl)
        p_Application.UserControl = blnUserControl
    End Property

    Public Property Get VBE()
        Set VBE = p_Application.VBE
    End Property

    Public Property Get Version()
        Version = p_Application.Version
    End Property

    Public Property Get Visible()
        Visible = p_Application.Visible
    End Property

    Public Property Let Visible(blnVisible)
        p_Application.Visible = blnVisible
    End Property

    Public Property Get WebServices()
        Set WebServices = p_Application.WebServices
    End Property


    ' Methods


    Public Function AccessError(intErrorNumber)
        AccessError = p_Application.AccessError(intErrorNumber)
    End Function

    Public Sub AddToFavorites()
        p_Application.AddToFavorites
    End Sub

    Public Function BuildCriteria(strField, intFieldType, strExpression)
        BuildCriteria = p_Application.BuildCriteria(strField, intFieldType, strExpression)
    End Function

    Public Sub CloseCurrentDatabase()
        p_Application.CloseCurrentDatabase
    End Sub

    Public Function CodeDb()
        Set CodeDb = p_Application.CodeDb()
    End Function

    Public Function ColumnHistory(strTableName, strColumnName, strQueryString)
        ColumnHistory = p_Application.ColumnHistory(strTableName, strColumnName, strQueryString)
    End Function

    Public Function CompactRepair(strSourceFile, strDestinationFile) ' Optional params: [LogFile As Boolean = False])
        CompactRepair = p_Application.CompactRepair(strSourceFile, strDestinationFile)
    End Function

    Public Sub ConvertAccessProject(strSourceFilename, strDestinationFilename, objDestinationFileFormat)
        p_Application.ConvertAccessProject strSourceFilename, strDestinationFilename, objDestinationFileFormat
    End Sub

    Public Sub CreateAccessProject(strFilepath) ' Optional param: [Connect])
        p_Application.CreateAccessProject strFilepath
    End Sub

    Public Function CreateAdditionalData()
        Set CreateAdditionalData = p_Application.CreateAdditionalData()
    End Function

    Public Function CreateControl(strFormName, objControlType) ' Optional params: [Section As AcSection = acDetail], [Parent], [ColumnName], [Left], [Top], [Width], [Height]) As Control
        Set CreateControl = p_Application.CreateControl(strFormName, objControlType)
    End Function

    Public Function CreateForm() ' Optional Params: [Database], [FormTemplate]) As Form
        Set CreateForm = p_Application.CreateForm()
    End Function

    Public Function CreateGroupLevel(strReportName, strExpression, intHeader, intFooter)
        CreateGroupLevel = p_Application.CreateGroupLevel(strReportName, strExpression, intHeader, intFooter)
    End Function

    Public Function CreateReport() ' Optional Params: [Database], [ReportTemplate]) As Report
        Set CreateReport = p_Application.CreateReport()
    End Function

    Public Function CreateReportControl(strReportName, objControlType) ' Optional params: [Section As AcSection = acDetail], [Parent], [ColumnName], [Left], [Top], [Width], [Height]) As Control
        Set CreateReportControl = p_Application.CreateReportControl(strReportName, objControlType)
    End Function

    Public Function CurrentDb()
        Set CurrentDb = p_Application.CurrentDb()
    End Function

    Public Function CurrentUser()
        CurrentUser = p_Application.CurrentUser
    End Function

    Public Function CurrentWebUser(objDisplayOption)
        Set CurrentWebUser = p_Application.CurrentWebUser
    End Function

    Public Function CurrentWebUserGroups(objDisplayOption)
        CurrentWebUserGroups = p_Application.CurrentWebUserGroups(objDisplayOption)
    End Function

    Public Function DAvg(strExpr, strDomain) ' Optional param: [Criteria])
        DAvg = p_Application.DAvg(strExpr, strDomain) 
    End Function

    Public Function DCount(strExpr, strDomain) ' Optional param: [Criteria])
        DCount = p_Application.DCount(strExpr, strDomain)
    End Function

    Public Sub DDEExecute(varChanNum, strCommand)
        p_Application.DDEExecute varChanNum, strCommand
    End Sub

    Public Function DDEInitiate(strApplication, strTopic)
        DDEInitiate = p_Application.DDEInitiate(strApplication, strTopic)
    End Function

    Public Sub DDEPoke(varChanNum, strItem, strData)
        p_Application.DDEPoke varChanNum, strItem, strData
    End Sub

    Public Function DDERequest(varChanNum, strItem)
        DDERequest = p_Application.DDERequest(varChanNum, strItem)
    End Function

    Public Sub DDETerminate(varChanNum)
        p_Application.DDETerminate varChanNum
    End Sub

    Public Sub DDETerminateAll()
        p_Application.DDETerminateAll
    End Sub

    Public Function DefaultWorkspaceClone()
        Set DefaultWorkspaceClone = p_Application.DefaultWorkspaceClone()
    End Function

    Public Sub DeleteControl(strFormName, strControlName)
        p_Application.DeleteControl strFormName, strControlName
    End Sub

    Public Sub DeleteReportControl(strReportName, strControlName)
        p_Application.DeleteReportControl strReportName, strControlName
    End Sub

    Public Function DFirst(strExpr, strDomain) ' Optional params: [Criteria]
        DFirst = p_Application.DFirst(strExpr, strDomain)
    End Function

    Public Sub DirtyObject(objObjectType, strObjectName)
        p_Application.DirtyObject objObjectType, strObjectName
    End Sub

    Public Function DLast(strExpr, strDomain) ' Optional params: [Criteria]
        DLast = p_Application.DLast(strExpr, strDomain)
    End Function

    Public Function DLookup(strExpr, strDomain) ' Optional params: [Criteria]
        DLookup = p_Application.DLookup(strExpr, strDomain)
    End Function

    Public Function DMax(strExpr, strDomain) ' Optional params: [Criteria]
        DMax = p_Application.DMax(strExpr, strDomain)
    End Function

    Public Function DMin(strExpr, strDomain) ' Optional params: [Criteria]
        DMin = p_Application.DMin(strExpr, strDomain)
    End Function

    Public Function DStDev(strExpr, strDomain) ' Optional params: [Criteria]
        DStDev = p_Application.DStDev(strExpr, strDomain)
    End Function

    Public Function DStDevP(strExpr, strDomain) ' Optional params: [Criteria]
        DStDevP = p_Application.DStDevP(strExpr, strDomain)
    End Function

    Public Function DSum(strExpr, strDomain) ' Optional params: [Criteria]
        DSum = p_Application.DSum(strExpr, strDomain)
    End Function

    Public Function DVar(strExpr, strDomain) ' Optional params: [Criteria]
        DVar = p_Application.DVar(strExpr, strDomain)
    End Function

    Public Function DVarP(strExpr, strDomain) ' Optional params: [Criteria]
        DVarP = p_Application.DVarP(strExpr, strDomain)
    End Function

    Public Sub Echo(intEchoOn) ' Optional params: [bstrStatusBarText As String])
        p_Application.Echo intEchoOn
    End Sub

    Public Function EuroConvert(dblNumber, strSourceCurrency, strTargetCurrency) ' Optional params: [FullPrecision], [TriangulationPrecision]) As Double
        EuroConvert = p_Application.EuroConvert(dblNumber, strSourceCurrency, strTargetCurrency)
    End Function

    Public Function Eval(strStringExpr)
        Eval = p_Application.Eval(strStringExpr)
    End Function

    Public Sub ExportNavigationPane(strPath)
        p_Application.ExportNavigationPane strPath
    End Sub

    Public Sub ExportXML(objObjectType, strDataSource) ' Optional params: [DataTarget As String], [SchemaTarget As String], [PresentationTarget As String], [ImageTarget As String], [Encoding As AcExportXMLEncoding = acUTF8], [OtherFlags As AcExportXMLOtherFlags], [WhereCondition As String], [AdditionalData])
        p_Application.ExportXML objObjectType, strDataSource
    End Sub

    Public Sub FollowHyperlink(strAddress) ' Optional params: [SubAddress As String], [NewWindow As Boolean = False], [AddHistory As Boolean = True], [ExtraInfo], [Method As MsoExtraInfoMethod], [HeaderInfo As String])
        p_Application.FollowHyperlink strAddress
    End Sub

    Public Function GetHiddenAttribute(objObjectType, strObjectName)
        GetHiddenAttribute = p_Application.GetHiddenAttribute(objObjectType, strObjectName)
    End Function

    Public Function GetOption(strOptionName)
        GetOption = p_Application.GetOption(strOptionName)
    End Function

    Public Function GUIDFromString(varString)
        GUIDFromString = p_Application.GUIDFromString(varString)
    End Function

    Public Function HtmlEncode(varPlainText) ' Optional params: [Length])
        HtmlEncode = p_Application.HtmlEncode(varPlainText)
    End Function

    Public Function hWndAccessApp()
        hWndAccessApp = p_Application.hWndAccessApp()
    End Function

    Public Function HyperlinkPart(varHyperlink) ' Optional params: [Part As AcHyperlinkPart = acDisplayedValue]) As String
        HyperlinkPart = p_Application.HyperlinkPart(varHyperlink)
    End Function

    Public Sub ImportNavigationPane(strPath) ' Optional params: [fAppendOnly As Boolean = False])
        p_Application.ImportNavigationPane strPath
    End Sub

    Public Sub ImportXML(strDataSource) ' Optional params: [ImportOptions As AcImportXMLOption = acStructureAndData])
        p_Application.ImportXML strDataSource
    End Sub

    Public Sub InstantiateTemplate(strPath)
        p_Application.InstantiateTemplate strPath
    End Sub

    Public Function IsCurrentWebUserInGroup(varGroupNameOrID)
        IsCurrentWebUserInGroup = p_Application.IsCurrentWebUserInGroup(varGroupNameOrID)
    End Function

    Public Sub LoadCustomUI(strCustomUIName, strCustomUIXML)
        p_Application.LoadCustomUI strCustomUIName, strCustomUIXML
    End Sub

    Public Sub LoadFromAXL(objObjectType, strObjectName, strFileName)
        p_Application.LoadFromAXL objObjectType, strObjectName, strFileName
    End Sub

    Public Function LoadPicture(strFileName)
        Set LoadPicture = p_Application.LoadPicture(strFileName)
    End Function

    Public Sub NewAccessProject(strFilepath) ' Optional params: [Connect]
        p_Application.NewAccessProject strFilepath
    End Sub

    Public Sub NewCurrentDatabase(strFilepath) ' Optional params: [FileFormat As AcNewDatabaseFormat = acNewDatabaseFormatUserDefault], [Template], [SiteAddress As String], [ListID As String])
        p_Application.NewCurrentDatabase strFilepath
    End Sub

    Public Function Nz(varValue) ' Optional params: [ValueIfNull]
        Nz = p_Application.Nz(varValue)
    End Function

    Public Sub OpenAccessProject(strFilepath) ' Optional params: [Exclusive As Boolean = False])
        p_Application.OpenAccessProject strFilepath
    End Sub

    Public Sub OpenCurrentDatabase(strFilepath) ' Optional params: [Exclusive As Boolean = False], [bstrPassword As String])
        p_Application.OpenCurrentDatabase strFilepath
    End Sub

    Public Function PlainText(varRichText) ' Optional params: [Length]) As String
        PlainText = p_Application.PlainText
    End Function

    Public Sub Quit() ' Optional params: [Option As AcQuitOption = acQuitSaveAll])
        p_Application.Quit
    End Sub

    Public Sub RefreshDatabaseWindow()
        p_Application.RefreshDatabaseWindow
    End Sub

    Public Sub RefreshTitleBar()
        p_Application.RefreshTitleBar
    End Sub

    Public Function Run(strProcedure) ' Optional params: [Arg1], [Arg2], [Arg3], [Arg4], [Arg5], [Arg6], [Arg7], [Arg8], [Arg9], [Arg10], [Arg11], [Arg12], [Arg13], [Arg14], [Arg15], [Arg16], [Arg17], [Arg18], [Arg19], [Arg20], [Arg21], [Arg22], [Arg23], [Arg24], [Arg25], [Arg26], [Arg27], [Arg28], [Arg29], [Arg30])
        Run = p_Application.Run
    End Function

    Public Sub RunCommand(objCommand)
        p_Application.RunCommand objCommand
    End Sub

    Public Sub SaveAsAXL(objObjectType, strObjectName, strFileName)
        p_Application.SaveAsAXL objObjectType, strObjectName, strFileName
    End Sub

    Public Sub SaveAsTemplate(strPath, strTitle, strIconPath, strCoreTable, strCategory) ' Optional params: [PreviewPath], [Description], [InstantiationForm], [ApplicationPart], [IncludeData], [Variation])
        p_Application.SaveAsTemplate strPath, strTitle, strIconPath, strCoreTable, strCategory
    End Sub 

    Public Sub SetDefaultWorkgroupFile(strPath)
        p_Application.SetDefaultWorkgroupFile strPath
    End Sub

    Public Sub SetHiddenAttribute(objObjectType, strObjectName, blnFHidden)
        p_Application.SetHiddenAttribute objObjectType, strObjectName, blnFHidden
    End Sub

    Public Sub SetOption(strOptionName, varSetting)
        p_Application.SetOption strOptionName, varSetting
    End Sub

    Public Function StringFromGUID(varGuid)
        StringFromGUID = p_Application.StringFromGUID(varGuid)
    End Function

    Public Function SysCmd(objAction) ' Optional params: [Argument2], [Argument3]
        SysCmd = p_Application.SysCmd(objAction)
    End Function

    Public Sub TransformXML(strDataSource, strTransformSource, strOutputTarget) ' Optional params: [WellFormedXMLOutput As Boolean = False], [ScriptOption As AcTransformXMLScriptOption = acPromptScript])
        p_Application.TransformXML strDataSource, strTransformSource, strOutputTarget
    End Sub

    Private Sub Class_Terminate()
        Set p_Application = Nothing
    End Sub
End Class

If WScript.ScriptName = "base_Access_Application.vbs" Then

End If