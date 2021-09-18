Option Explicit

Class base_App_objAccess
    Private p_objAccess

    Private Sub Class_Initialize()
        Set p_objAccess = CreateObject("Access.Application")
    End Sub


    ' Properties


    Public Property Get Application()
        Set Application = p_objAccess.Application
    End Property

    Public Property Get Assistance()
        Set Assistance = p_objAccess.Assistance
    End Property

    Public Property Get AutoCorrect()
        Set AutoCorrect = p_objAccess.AutoCorrect
    End Property

    Public Property Get AutomationSecurity()
        Set AutomationSecurity = p_objAccess.AutomationSecurity
    End Property

    Public Property Get BrokenReference()
        BrokenReference = p_objAccess.BrokenReference
    End Property

    Public Property Get Build()
        Build = p_objAccess.Build
    End Property

    Public Property Get CodeContextObject()
        Set CodeContextObject = p_objAccess.CodeContextObject
    End Property

    Public Property Get CodeData()
        Set CodeData = p_objAccess.CodeData
    End Property

    Public Property Get CodeProject()
        Set CodeProject = p_objAccess.CodeProject
    End Property

    Public Property Get COMAddIns()
        Set COMAddIns = p_objAccess.COMAddIns
    End Property

    Public Property Get CommandBars()
        Set CommandBars = p_objAccess.CommandBars
    End Property

    Public Property Get CurrentData()
        Set CurrentData = p_objAccess.CurrentData
    End Property

    Public Property Get CurrentObjectName()
        CurrentObjectName = p_objAccess.CurrentObjectName
    End Property

    Public Property Get CurrentObjectType()
        Set CurrentObjectType = p_objAccess.CurrentObjectType
    End Property

    Public Property Get CurrentProject()
        Set CurrentProject = p_objAccess.CurrentProject
    End Property

    Public Property Get DBEngine()
        Set DBEngine = p_objAccess.DBEngine
    End Property

    Public Property Get DoCmd()
        Set DoCmd = p_objAccess.DoCmd
    End Property

    Public Property Get FeatureInstall()
        Set FeatureInstall = p_objAccess.FeatureInstall
    End Property

    Public Property Set FeatureInstall(objMsoFeatureInstall)
        Set p_objAccess.FeatureInstall = objMsoFeatureInstall
    End Property

    Public Property Get FileDialog()
        Set FileDialog = p_objAccess.FileDialog
    End Property

    Public Property Get Forms()
        Set Forms = p_objAccess.Forms
    End Property

    Public Property Get IsCompiled()
        IsCompiled = p_objAccess.IsCompiled
    End Property

    Public Property Get LanguageSettings()
        Set LanguageSettings = p_objAccess.LanguageSettings
    End Property

    Public Property Get MacroError()
        Set MacroError = p_objAccess.MacroError
    End Property

    Public Property Get MenuBar()
        MenuBar = p_objAccess.MenuBar
    End Property

    Public Property Let MenuBar(strMenuBar)
        p_objAccess.MenuBar = strMenuBar
    End Property

    Public Property Get Modules()
        Set Modules = p_objAccess.Modules
    End Property

    Public Property Get Name()
        Name = p_objAccess.Name
    End Property

    Public Property Get NewFileTaskPane()
        Set NewFileTaskPane = p_objAccess.NewFileTaskPane
    End Property

    Public Property Get Parent()
        Set Parent = p_objAccess.Parent
    End Property

    Public Property Get Printer()
        Set Printer = p_objAccess.Printer
    End Property

    Public Property Set Printer(objPrinter)
        Set p_objAccess.Printer = objPrinter
    End Property

    Public Property Get Printers()
        Set Printers = p_objAccess.Printers
    End Property

    Public Property Get ProductCode()
        ProductCode = p_objAccess.ProductCode
    End Property

    Public Property Get References()
        Set References = p_objAccess.References
    End Property

    Public Property Get Reports()
        Set Reports = p_objAccess.Reports
    End Property

    Public Property Get ReturnVars()
        Set ReturnVars = p_objAccess.ReturnVars
    End Property

    Public Property Get Screen()
        Set Screen = p_objAccess.Screen
    End Property

    Public Property Get ShortcutMenuBar()
        ShortcutMenuBar = p_objAccess.ShortcutMenuBar
    End Property

    Public Property Let ShortcutMenuBar(strShortcutMenubar)
        p_objAccess.ShortcutMenuBar = strShortcutMenubar
    End Property

    Public Property Get TempVars()
        Set TempVars = p_objAccess.TempVars
    End Property

    Public Property Get UserControl()
        UserControl = p_objAccess.UserControl
    End Property

    Public Property Let UserControl(blnUserControl)
        p_objAccess.UserControl = blnUserControl
    End Property

    Public Property Get VBE()
        Set VBE = p_objAccess.VBE
    End Property

    Public Property Get Version()
        Version = p_objAccess.Version
    End Property

    Public Property Get Visible()
        Visible = p_objAccess.Visible
    End Property

    Public Property Let Visible(blnVisible)
        p_objAccess.Visible = blnVisible
    End Property

    Public Property Get WebServices()
        Set WebServices = p_objAccess.WebServices
    End Property


    ' Methods


    Public Function AccessError(intErrorNumber)
        AccessError = p_objAccess.AccessError(intErrorNumber)
    End Function

    Public Sub AddToFavorites()
        p_objAccess.AddToFavorites
    End Sub

    Public Function BuildCriteria(strField, intFieldType, strExpression)
        BuildCriteria = p_objAccess.BuildCriteria(strField, intFieldType, strExpression)
    End Function

    Public Sub CloseCurrentDatabase()
        p_objAccess.CloseCurrentDatabase
    End Sub

    Public Function CodeDb()
        Set CodeDb = p_objAccess.CodeDb()
    End Function

    Public Function ColumnHistory(strTableName, strColumnName, strQueryString)
        ColumnHistory = p_objAccess.ColumnHistory(strTableName, strColumnName, strQueryString)
    End Function

    Public Function CompactRepair(strSourceFile, strDestinationFile) ' Optional params: [LogFile As Boolean = False])
        CompactRepair = p_objAccess.CompactRepair(strSourceFile, strDestinationFile)
    End Function

    Public Sub ConvertAccessProject(strSourceFilename, strDestinationFilename, objDestinationFileFormat)
        p_objAccess.ConvertAccessProject strSourceFilename, strDestinationFilename, objDestinationFileFormat
    End Sub

    Public Sub CreateAccessProject(strFilepath) ' Optional param: [Connect])
        p_objAccess.CreateAccessProject strFilepath
    End Sub

    Public Function CreateAdditionalData()
        Set CreateAdditionalData = p_objAccess.CreateAdditionalData()
    End Function

    Public Function CreateControl(strFormName, objControlType) ' Optional params: [Section As AcSection = acDetail], [Parent], [ColumnName], [Left], [Top], [Width], [Height]) As Control
        Set CreateControl = p_objAccess.CreateControl(strFormName, objControlType)
    End Function

    Public Function CreateForm() ' Optional Params: [Database], [FormTemplate]) As Form
        Set CreateForm = p_objAccess.CreateForm()
    End Function

    Public Function CreateGroupLevel(strReportName, strExpression, intHeader, intFooter)
        CreateGroupLevel = p_objAccess.CreateGroupLevel(strReportName, strExpression, intHeader, intFooter)
    End Function

    Public Function CreateReport() ' Optional Params: [Database], [ReportTemplate]) As Report
        Set CreateReport = p_objAccess.CreateReport()
    End Function

    Public Function CreateReportControl(strReportName, objControlType) ' Optional params: [Section As AcSection = acDetail], [Parent], [ColumnName], [Left], [Top], [Width], [Height]) As Control
        Set CreateReportControl = p_objAccess.CreateReportControl(strReportName, objControlType)
    End Function

    Public Function CurrentDb()
        Set CurrentDb = p_objAccess.CurrentDb()
    End Function

    Public Function CurrentUser()
        CurrentUser = p_objAccess.CurrentUser
    End Function

    Public Function CurrentWebUser(objDisplayOption)
        Set CurrentWebUser = p_objAccess.CurrentWebUser
    End Function

    Public Function CurrentWebUserGroups(objDisplayOption)
        CurrentWebUserGroups = p_objAccess.CurrentWebUserGroups(objDisplayOption)
    End Function

    Public Function DAvg(strExpr, strDomain) ' Optional param: [Criteria])
        DAvg = p_objAccess.DAvg(strExpr, strDomain) 
    End Function

    Public Function DCount(strExpr, strDomain) ' Optional param: [Criteria])
        DCount = p_objAccess.DCount(strExpr, strDomain)
    End Function

    Public Sub DDEExecute(varChanNum, strCommand)
        p_objAccess.DDEExecute varChanNum, strCommand
    End Sub

    Public Function DDEInitiate(strApplication, strTopic)
        DDEInitiate = p_objAccess.DDEInitiate(strApplication, strTopic)
    End Function

    Public Sub DDEPoke(varChanNum, strItem, strData)
        p_objAccess.DDEPoke varChanNum, strItem, strData
    End Sub

    Public Function DDERequest(varChanNum, strItem)
        DDERequest = p_objAccess.DDERequest(varChanNum, strItem)
    End Function

    Public Sub DDETerminate(varChanNum)
        p_objAccess.DDETerminate varChanNum
    End Sub

    Public Sub DDETerminateAll()
        p_objAccess.DDETerminateAll
    End Sub

    Public Function DefaultWorkspaceClone()
        Set DefaultWorkspaceClone = p_objAccess.DefaultWorkspaceClone()
    End Function

    Public Sub DeleteControl(strFormName, strControlName)
        p_objAccess.DeleteControl strFormName, strControlName
    End Sub

    Public Sub DeleteReportControl(strReportName, strControlName)
        p_objAccess.DeleteReportControl strReportName, strControlName
    End Sub

    Public Function DFirst(strExpr, strDomain) ' Optional params: [Criteria]
        DFirst = p_objAccess.DFirst(strExpr, strDomain)
    End Function

    Public Sub DirtyObject(objObjectType, strObjectName)
        p_objAccess.DirtyObject objObjectType, strObjectName
    End Sub

    Public Function DLast(strExpr, strDomain) ' Optional params: [Criteria]
        DLast = p_objAccess.DLast(strExpr, strDomain)
    End Function

    Public Function DLookup(strExpr, strDomain) ' Optional params: [Criteria]
        DLookup = p_objAccess.DLookup(strExpr, strDomain)
    End Function

    Public Function DMax(strExpr, strDomain) ' Optional params: [Criteria]
        DMax = p_objAccess.DMax(strExpr, strDomain)
    End Function

    Public Function DMin(strExpr, strDomain) ' Optional params: [Criteria]
        DMin = p_objAccess.DMin(strExpr, strDomain)
    End Function

    Public Function DStDev(strExpr, strDomain) ' Optional params: [Criteria]
        DStDev = p_objAccess.DStDev(strExpr, strDomain)
    End Function

    Public Function DStDevP(strExpr, strDomain) ' Optional params: [Criteria]
        DStDevP = p_objAccess.DStDevP(strExpr, strDomain)
    End Function

    Public Function DSum(strExpr, strDomain) ' Optional params: [Criteria]
        DSum = p_objAccess.DSum(strExpr, strDomain)
    End Function

    Public Function DVar(strExpr, strDomain) ' Optional params: [Criteria]
        DVar = p_objAccess.DVar(strExpr, strDomain)
    End Function

    Public Function DVarP(strExpr, strDomain) ' Optional params: [Criteria]
        DVarP = p_objAccess.DVarP(strExpr, strDomain)
    End Function

    Public Sub Echo(intEchoOn) ' Optional params: [bstrStatusBarText As String])
        p_objAccess.Echo intEchoOn
    End Sub

    Public Function EuroConvert(dblNumber, strSourceCurrency, strTargetCurrency) ' Optional params: [FullPrecision], [TriangulationPrecision]) As Double
        EuroConvert = p_objAccess.EuroConvert(dblNumber, strSourceCurrency, strTargetCurrency)
    End Function

    Public Function Eval(strStringExpr)
        Eval = p_objAccess.Eval(strStringExpr)
    End Function

    Public Sub ExportNavigationPane(strPath)
        p_objAccess.ExportNavigationPane strPath
    End Sub

    Public Sub ExportXML(objObjectType, strDataSource) ' Optional params: [DataTarget As String], [SchemaTarget As String], [PresentationTarget As String], [ImageTarget As String], [Encoding As AcExportXMLEncoding = acUTF8], [OtherFlags As AcExportXMLOtherFlags], [WhereCondition As String], [AdditionalData])
        p_objAccess.ExportXML objObjectType, strDataSource
    End Sub

    Public Sub FollowHyperlink(strAddress) ' Optional params: [SubAddress As String], [NewWindow As Boolean = False], [AddHistory As Boolean = True], [ExtraInfo], [Method As MsoExtraInfoMethod], [HeaderInfo As String])
        p_objAccess.FollowHyperlink strAddress
    End Sub

    Public Function GetHiddenAttribute(objObjectType, strObjectName)
        GetHiddenAttribute = p_objAccess.GetHiddenAttribute(objObjectType, strObjectName)
    End Function

    Public Function GetOption(strOptionName)
        GetOption = p_objAccess.GetOption(strOptionName)
    End Function

    Public Function GUIDFromString(varString)
        GUIDFromString = p_objAccess.GUIDFromString(varString)
    End Function

    Public Function HtmlEncode(varPlainText) ' Optional params: [Length])
        HtmlEncode = p_objAccess.HtmlEncode(varPlainText)
    End Function

    Public Function hWndAccessApp()
        hWndAccessApp = p_objAccess.hWndAccessApp()
    End Function

    Public Function HyperlinkPart(varHyperlink) ' Optional params: [Part As AcHyperlinkPart = acDisplayedValue]) As String
        HyperlinkPart = p_objAccess.HyperlinkPart(varHyperlink)
    End Function

    Public Sub ImportNavigationPane(strPath) ' Optional params: [fAppendOnly As Boolean = False])
        p_objAccess.ImportNavigationPane strPath
    End Sub

    Public Sub ImportXML(strDataSource) ' Optional params: [ImportOptions As AcImportXMLOption = acStructureAndData])
        p_objAccess.ImportXML strDataSource
    End Sub

    Public Sub InstantiateTemplate(strPath)
        p_objAccess.InstantiateTemplate strPath
    End Sub

    Public Function IsCurrentWebUserInGroup(varGroupNameOrID)
        IsCurrentWebUserInGroup = p_objAccess.IsCurrentWebUserInGroup(varGroupNameOrID)
    End Function

    Public Sub LoadCustomUI(strCustomUIName, strCustomUIXML)
        p_objAccess.LoadCustomUI strCustomUIName, strCustomUIXML
    End Sub

    Public Sub LoadFromAXL(objObjectType, strObjectName, strFileName)
        p_objAccess.LoadFromAXL objObjectType, strObjectName, strFileName
    End Sub

    Public Function LoadPicture(strFileName)
        Set LoadPicture = p_objAccess.LoadPicture(strFileName)
    End Function

    Public Sub NewAccessProject(strFilepath) ' Optional params: [Connect]
        p_objAccess.NewAccessProject strFilepath
    End Sub

    Public Sub NewCurrentDatabase(strFilepath) ' Optional params: [FileFormat As AcNewDatabaseFormat = acNewDatabaseFormatUserDefault], [Template], [SiteAddress As String], [ListID As String])
        p_objAccess.NewCurrentDatabase strFilepath
    End Sub

    Public Function Nz(varValue) ' Optional params: [ValueIfNull]
        Nz = p_objAccess.Nz(varValue)
    End Function

    Public Sub OpenAccessProject(strFilepath) ' Optional params: [Exclusive As Boolean = False])
        p_objAccess.OpenAccessProject strFilepath
    End Sub

    Public Sub OpenCurrentDatabase(strFilepath) ' Optional params: [Exclusive As Boolean = False], [bstrPassword As String])
        p_objAccess.OpenCurrentDatabase strFilepath
    End Sub

    Public Function PlainText(varRichText) ' Optional params: [Length]) As String
        PlainText = p_objAccess.PlainText
    End Function

    Public Sub Quit() ' Optional params: [Option As AcQuitOption = acQuitSaveAll])
        p_objAccess.Quit
    End Sub

    Public Sub RefreshDatabaseWindow()
        p_objAccess.RefreshDatabaseWindow
    End Sub

    Public Sub RefreshTitleBar()
        p_objAccess.RefreshTitleBar
    End Sub

    Public Function Run(strProcedure) ' Optional params: [Arg1], [Arg2], [Arg3], [Arg4], [Arg5], [Arg6], [Arg7], [Arg8], [Arg9], [Arg10], [Arg11], [Arg12], [Arg13], [Arg14], [Arg15], [Arg16], [Arg17], [Arg18], [Arg19], [Arg20], [Arg21], [Arg22], [Arg23], [Arg24], [Arg25], [Arg26], [Arg27], [Arg28], [Arg29], [Arg30])
        Run = p_objAccess.Run
    End Function

    Public Sub RunCommand(objCommand)
        p_objAccess.RunCommand objCommand
    End Sub

    Public Sub SaveAsAXL(objObjectType, strObjectName, strFileName)
        p_objAccess.SaveAsAXL objObjectType, strObjectName, strFileName
    End Sub

    Public Sub SaveAsTemplate(strPath, strTitle, strIconPath, strCoreTable, strCategory) ' Optional params: [PreviewPath], [Description], [InstantiationForm], [ApplicationPart], [IncludeData], [Variation])
        p_objAccess.SaveAsTemplate strPath, strTitle, strIconPath, strCoreTable, strCategory
    End Sub 

    Public Sub SetDefaultWorkgroupFile(strPath)
        p_objAccess.SetDefaultWorkgroupFile strPath
    End Sub

    Public Sub SetHiddenAttribute(objObjectType, strObjectName, blnFHidden)
        p_objAccess.SetHiddenAttribute objObjectType, strObjectName, blnFHidden
    End Sub

    Public Sub SetOption(strOptionName, varSetting)
        p_objAccess.SetOption strOptionName, varSetting
    End Sub

    Public Function StringFromGUID(varGuid)
        StringFromGUID = p_objAccess.StringFromGUID(varGuid)
    End Function

    Public Function SysCmd(objAction) ' Optional params: [Argument2], [Argument3]
        SysCmd = p_objAccess.SysCmd(objAction)
    End Function

    Public Sub TransformXML(strDataSource, strTransformSource, strOutputTarget) ' Optional params: [WellFormedXMLOutput As Boolean = False], [ScriptOption As AcTransformXMLScriptOption = acPromptScript])
        p_objAccess.TransformXML strDataSource, strTransformSource, strOutputTarget
    End Sub

    Private Sub Class_Terminate()
        Set p_objAccess = Nothing
    End Sub
End Class

If WScript.ScriptName = "base_App_objAccess.vbs" Then

End If