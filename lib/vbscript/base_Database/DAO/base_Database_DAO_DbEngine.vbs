Option Explicit

Class base_Database_DAO_DbEngine
	Private p_DbEngine

	Private Sub Class_Initialize()
		Set p_DbEngine = CreateObject("DAO.DBEngine.36")
	End Sub


	' Properties


	Public Property Get DefaultPassword()
		DefaultPassword = p_DbEngine.DefaultPassword
	End Property

	Public Property Let DefaultPassword(strDefaultPassword)
		p_DbEngine.DefaultPassword = strDefaultPassword
	End Property

	Public Property Get DefaultType()
		DefaultType = p_DbEngine.DefaultType
	End Property

	Public Property Let DefaultType(lngDefaultType)
		p_DbEngine.DefaultType = lngDefaultType
	End Property

	Public Property Get DefaultUser()
		DefaultUser = p_DbEngine.DefaultUser
	End Property

	Public Property Let DefaultUser(strDefaultUser)
		p_DbEngine.DefaultUser = strDefaultUser
	End Property

	Public Property Get Errors()
		Set Errors = p_DbEngine.Errors
	End Property

	Public Property Get IniPath()
		IniPath = p_DbEngine.IniPath
	End Property

	Public Property Let IniPath(strIniPath)
		p_DbEngine.IniPath =  strIniPath
	End Property

	Public Property Get LoginTimeout()
		LoginTimeout = p_DbEngine.LoginTimeout
	End Property

	Public Property Let LoginTimeout(intLoginTimeout)
		p_DbEngine.LoginTimeout = intLoginTimeout
	End Property

	Public Property Get Properties()
		Set Properties = p_DbEngine.Properties
	End Property

	Public Property Get SystemDB()
		SystemDB = p_DbEngine.SystemDB
	End Property

	Public Property Let SystemDB(strSystemDB)
		p_DbEngine.SystemDB = strSystemDB
	End Property

	Public Property Get Version()
		Version = p_DbEngine.Version
	End Property

	Public Default Property Get Workspaces()
		Set Workspaces = p_DbEngine.Workspaces
	End Property


	' Methods


	Public Sub BeginTrans()
		p_DbEngine.BeginTrans
	End Sub

	Public Sub CommitTrans() ' Optional params: [Option As Long]
		p_DbEngine.CommitTrans
	End Sub

	Public Sub CompactDatabase(strSrcName, strDstName) ' Optional params: [DstLocale], [Options], [SrcLocale]
		p_DbEngine.CompactDatabase strSrcName, strDstName
	End Sub

	Public Function CreateDatabase(strName, strLocale) ' Optional params: [Option]
		Set CreateDatabase = p_DbEngine.CreateDatabase(strName, strLocale)
	End Function

	Public Function CreateWorkspace(strName, strUserName, strPassword) ' Optional params: [UseType]
		Set CreateWorkspace = p_DbEngine.CreateWorkspace(strName, strUserName, strPassword)
	End Function

	Public Sub Idle() ' Optional params: [Action]
		p_DbEngine.Idle
	End Sub

	Public Function OpenConnection(strName) ' Optional params: [Options], [ReadOnly], [Connect]
		Set OpenConnection = p_DbEngine.OpenConnection(strName)
	End Function

	Public Function OpenDatabase(strName) ' Optional params: [Options], [ReadOnly], [Connect]
		Set OpenDatabase = p_DbEngine.OpenDatabase(strName)
	End Function

	Public Sub RegisterDatabase(strDsn, strDriver, blnSilent, strAttributes)
		p_DbEngine.RegisterDatabase strDsn, strDriver, blnSilent, strAttributes
	End Sub

	Public Sub Rollback()
		p_DbEngine.Rollback
	End Sub

	Public Sub SetOption(lngOption, varValue)
		p_DbEngine.SetOption lngOption, varValue
	End Sub

	Private Sub Class_Terminate()
		Set p_DbEngine = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Database_DAO_DbEngine.vbs" Then
	Dim objDbEngine

	Set objDbEngine = New base_Database_DAO_DbEngine

	WScript.Echo objDbEngine.Version
End If
