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


    Public Function AccessError(ErrorNumber)
        AccessError = p_Application.AccessError
    End Function

    Public Sub AddToFavorites()
        p_Application.AddToFavorites
    End Sub

    Public Function BuildCriteria(Field As String, FieldType As Integer, Expression As String) As String
        BuildCriteria = p_Application.BuildCriteria
    End Function

    Public Sub CloseCurrentDatabase()
        p_Application.CloseCurrentDatabase
    End Sub

    Public Function CodeDb() As Database
        CodeDb = p_Application.CodeDb
    End Function

    Public Function ColumnHistory(TableName As String, ColumnName As String, queryString As String) As String
        ColumnHistory = p_Application.ColumnHistory
    End Function

    Public Function CompactRepair(SourceFile As String, DestinationFile As String, [LogFile As Boolean = False]) As Boolean
        CompactRepair = p_Application.CompactRepair
    End Function

    Public Sub ConvertAccessProject(SourceFilename As String, DestinationFilename As String, DestinationFileFormat As AcFileFormat)
        p_Application.
    End Sub

    Public Sub CreateAccessProject(filepath As String, [Connect])
        p_Application.
    End Sub

    Public Function CreateAdditionalData() As AdditionalData
        CreateAdditionalData = p_Application.CreateAdditionalData
    End Function

    Public Function CreateControl(FormName As String, ControlType As AcControlType, [Section As AcSection = acDetail], [Parent], [ColumnName], [Left], [Top], [Width], [Height]) As Control
        CreateControl = p_Application.CreateControl
    End Function

    Public Function CreateForm([Database], [FormTemplate]) As Form
        CreateForm = p_Application.CreateForm
    End Function

    Public Function CreateGroupLevel(ReportName As String, Expression As String, Header As Integer, Footer As Integer) As Long
        CreateGroupLevel = p_Application.CreateGroupLevel
    End Function

    Public Function CreateReport([Database], [ReportTemplate]) As Report
        CreateReport = p_Application.CreateReport
    End Function

    Public Function CreateReportControl(ReportName As String, ControlType As AcControlType, [Section As AcSection = acDetail], [Parent], [ColumnName], [Left], [Top], [Width], [Height]) As Control
        CreateReportControl = p_Application.CreateReportControl
    End Function

    Public Function CurrentDb() As Database
        CurrentDb = p_Application.CurrentDb
    End Function

    Public Function CurrentUser() As String
        CurrentUser = p_Application.CurrentUser
    End Function

    Public Function CurrentWebUser(DisplayOption As AcWebUserDisplay)
        CurrentWebUser = p_Application.CurrentWebUser
    End Function

    Public Function CurrentWebUserGroups(DisplayOption As AcWebUserGroupsDisplay)
        CurrentWebUserGroups = p_Application.CurrentWebUserGroups
    End Function

    Public Function DAvg(Expr As String, Domain As String, [Criteria])
        DAvg = p_Application.DAvg
    End Function

    Public Function DCount(Expr As String, Domain As String, [Criteria])
        DCount = p_Application.DCount
    End Function

    Public Sub DDEExecute(ChanNum, Command As String)
        p_Application.DDEExecute
    End Sub

    Public Function DDEInitiate(Application As String, Topic As String)
        DDEInitiate = p_Application.DDEInitiate
    End Function

    Public Sub DDEPoke(ChanNum, Item As String, Data As String)
        p_Application.DDEPoke
    End Sub

    Public Function DDERequest(ChanNum, Item As String) As String
        DDERequest = p_Application.DDERequest
    End Function

    Public Sub DDETerminate(ChanNum)
        p_Application.DDETerminate
    End Sub

    Public Sub DDETerminateAll()
        p_Application.DDETerminateAll
    End Sub

    Public Function DefaultWorkspaceClone() As Workspace
        DefaultWorkspaceClone = p_Application.DefaultWorkspaceClone
    End Function

    Public Sub DeleteControl(FormName As String, ControlName As String)
        p_Application.DeleteControl
    End Sub

    Public Sub DeleteReportControl(ReportName As String, ControlName As String)
        p_Application.DeleteReportControl
    End Sub

    Public Function DFirst(Expr As String, Domain As String, [Criteria])
        DFirst = p_Application.DFirst
    End Function

    Public Sub DirtyObject(ObjectType As AcObjectType, ObjectName As String)
        p_Application.DirtyObject
    End Sub

    Public Function DLast(Expr As String, Domain As String, [Criteria])
        DLast = p_Application.DLast
    End Function

    Public Function DLookup(Expr As String, Domain As String, [Criteria])
        DLookup = p_Application.DLookup
    End Function

    Public Function DMax(Expr As String, Domain As String, [Criteria])
        DMax = p_Application.DMax
    End Function

    Public Function DMin(Expr As String, Domain As String, [Criteria])
        DMin = p_Application.DMin
    End Function

    Public Function DStDev(Expr As String, Domain As String, [Criteria])
        DStDev = p_Application.DStDev
    End Function

    Public Function DStDevP(Expr As String, Domain As String, [Criteria])
        DStDevP = p_Application.DStDevP
    End Function

    Public Function DSum(Expr As String, Domain As String, [Criteria])
        DSum = p_Application.DSum
    End Function

    Public Function DVar(Expr As String, Domain As String, [Criteria])
        DVar = p_Application.DVar
    End Function

    Public Function DVarP(Expr As String, Domain As String, [Criteria])
        DVarP = p_Application.DVarP
    End Function

    Public Sub Echo(EchoOn As Integer, [bstrStatusBarText As String])
        p_Application.Echo
    End Sub

    Public Function EuroConvert(Number As Double, SourceCurrency As String, TargetCurrency As String, [FullPrecision], [TriangulationPrecision]) As Double
        EuroConvert = p_Application.EuroConvert
    End Function

    Public Function Eval(StringExpr As String)
        Eval = p_Application.Eval
    End Function

    Public Sub ExportNavigationPane(Path As String)
        p_Application.ExportNavigationPane
    End Sub

    Public Sub ExportXML(ObjectType As AcExportXMLObjectType, DataSource As String, [DataTarget As String], [SchemaTarget As String], [PresentationTarget As String], [ImageTarget As String], [Encoding As AcExportXMLEncoding = acUTF8], [OtherFlags As AcExportXMLOtherFlags], [WhereCondition As String], [AdditionalData])
        p_Application.ExportXML
    End Sub

    Public Sub FollowHyperlink(Address As String, [SubAddress As String], [NewWindow As Boolean = False], [AddHistory As Boolean = True], [ExtraInfo], [Method As MsoExtraInfoMethod], [HeaderInfo As String])
        p_Application.FollowHyperlink
    End Sub

    Public Function GetHiddenAttribute(ObjectType As AcObjectType, ObjectName As String) As Boolean
        GetHiddenAttribute = p_Application.GetHiddenAttribute
    End Function

    Public Function GetOption(OptionName As String)
        GetOption = p_Application.GetOption
    End Function

    Public Function GUIDFromString(String)
        GUIDFromString = p_Application.GUIDFromString
    End Function

    Public Function HtmlEncode(PlainText, [Length]) As String
        HtmlEncode = p_Application.HtmlEncode
    End Function

    Public Function hWndAccessApp() As Long
        hWndAccessApp = p_Application.hWndAccessApp
    End Function

    Public Function HyperlinkPart(Hyperlink, [Part As AcHyperlinkPart = acDisplayedValue]) As String
        HyperlinkPart = p_Application.HyperlinkPart
    End Function

    Public Sub ImportNavigationPane(Path As String, [fAppendOnly As Boolean = False])
        p_Application.ImportNavigationPane
    End Sub

    Public Sub ImportXML(DataSource As String, [ImportOptions As AcImportXMLOption = acStructureAndData])
        p_Application.ImportXML
    End Sub

    Public Sub InstantiateTemplate(Path As String)
        p_Application.InstantiateTemplate
    End Sub

    Public Function IsCurrentWebUserInGroup(GroupNameOrID) As Boolean
        IsCurrentWebUserInGroup = p_Application.IsCurrentWebUserInGroup
    End Function

    Public Sub LoadCustomUI(CustomUIName As String, CustomUIXML As String)
        p_Application.LoadCustomUI
    End Sub

    Public Sub LoadFromAXL(ObjectType As AcObjectType, ObjectName As String, FileName As String)
        p_Application.LoadFromAXL
    End Sub

    Public Function LoadPicture(FileName As String) As Object
        LoadPicture = p_Application.LoadPicture
    End Function

    Public Sub NewAccessProject(filepath As String, [Connect])
        p_Application.NewAccessProject
    End Sub

    Public Sub NewCurrentDatabase(filepath As String, [FileFormat As AcNewDatabaseFormat = acNewDatabaseFormatUserDefault], [Template], [SiteAddress As String], [ListID As String])
        p_Application.NewCurrentDatabase
    End Sub

    Public Function Nz(Value, [ValueIfNull])
        Nz = p_Application.Nz
    End Function

    Public Sub OpenAccessProject(filepath As String, [Exclusive As Boolean = False])
        p_Application.OpenAccessProject
    End Sub

    Public Sub OpenCurrentDatabase(filepath As String, [Exclusive As Boolean = False], [bstrPassword As String])
        p_Application.OpenCurrentDatabase
    End Sub

    Public Function PlainText(RichText, [Length]) As String
        PlainText = p_Application.PlainText
    End Function

    Public Sub Quit([Option As AcQuitOption = acQuitSaveAll])
        p_Application.Quit
    End Sub

    Public Sub RefreshDatabaseWindow()
        p_Application.RefreshDatabaseWindow
    End Sub

    Public Sub RefreshTitleBar()
        p_Application.RefreshTitleBar
    End Sub

    Public Function Run(Procedure As String, [Arg1], [Arg2], [Arg3], [Arg4], [Arg5], [Arg6], [Arg7], [Arg8], [Arg9], [Arg10], [Arg11], [Arg12], [Arg13], [Arg14], [Arg15], [Arg16], [Arg17], [Arg18], [Arg19], [Arg20], [Arg21], [Arg22], [Arg23], [Arg24], [Arg25], [Arg26], [Arg27], [Arg28], [Arg29], [Arg30])
        Run = p_Application.Run
    End Function

    Public Sub RunCommand(Command As AcCommand)
        p_Application.RunCommand
    End Sub

    Public Sub SaveAsAXL(ObjectType As AcObjectType, ObjectName As String, FileName As String)
        p_Application.SaveAsAXL
    End Sub

    Public Sub SaveAsTemplate(Path As String, Title As String, IconPath As String, CoreTable As String, Category As String, [PreviewPath], [Description], [InstantiationForm], [ApplicationPart], [IncludeData], [Variation])
        p_Application.SaveAsTemplate
    End Sub

    Public Sub SetDefaultWorkgroupFile(Path As String)
        p_Application.SetDefaultWorkgroupFile
    End Sub

    Public Sub SetHiddenAttribute(ObjectType As AcObjectType, ObjectName As String, fHidden As Boolean)
        p_Application.SetHiddenAttribute
    End Sub

    Public Sub SetOption(OptionName As String, Setting)
        p_Application.SetOption
    End Sub

    Public Function StringFromGUID(Guid)
        StringFromGUID = p_Application.StringFromGUID
    End Function

    Public Function SysCmd(Action As AcSysCmdAction, [Argument2], [Argument3])
        SysCmd = p_Application.SysCmd
    End Function

    Public Sub TransformXML(DataSource As String, TransformSource As String, OutputTarget As String, [WellFormedXMLOutput As Boolean = False], [ScriptOption As AcTransformXMLScriptOption = acPromptScript])
        p_Application.TransformXML
    End Sub

    Private Sub Class_Terminate()
        Set p_Application = Nothing
    End Sub
End Class

If WScript.ScriptName = "base_Access_Application.vbs" Then

End If
