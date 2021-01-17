Option Explicit

Private pScheme, _
	pSchemeName, _
	pHostname, _
	pSubdomain, _
	pDomain, _
	pTLD, _
	pUsername, _
	pPassword, _
	pUserinfo, _
	pAuthority, _
	pPort, _
	pPath, _
	pDirectory, _
	pFilename, _
	pSuffix, _
	pQuery, _
	pFragment

Private pUriRegex, _
	pUriParserHost, _
	pUriParser, _
	pUriCoder

Private Sub Class_Initialize()
	pScheme = ""
	pSchemeName = ""
	pHostname = ""
	pSubdomain = ""
	pDomain = ""
	pTLD = ""
	pUsername = ""
	pPassword = ""
	pUserinfo = ""
	pAuthority = ""
	pPort = ""
	pPath = ""
	pDirectory = ""
	pFilename = ""
	pSuffix = ""
	pQuery = ""
	pFragment = ""

	Set pUriRegex = New RegExp
	Set pUriParserHost = CreateObject("HTMLFile")
	Set pUriParser = pUriParserHost.createElement("a")
	Set pUriCoder = New v_Script

	With pUriCoder
		.Language = "JScript"
		.AddCode "function encode(uri) { return encodeURIComponent(uri); }"
		.AddCode "function decode(uri) { return decodeURIComponent(uri); }"
	End With
End Sub


' Properties


Public Property Get Scheme()
	Scheme = pScheme
End Property

Public Property Let Scheme(strScheme)
	Me.FromString strScheme & "://" & Me.Authority & Me.Resource
End Property

Public Property Get SchemeName()
	SchemeName = pSchemeName
End Property	

Public Property Get Protocol()
	Protocol = pScheme
End Property

Public Property Let Protocol(strProtocol)
	Me.FromString strProtocol & "://" & Me.Authority & Me.Resource
End Property

Public Property Get ProtocolName()
	ProtocolName = pSchemeName
End Property

Public Property Get Host()
	Host = pHostname
End Property

Public Property Let Host(strHost)
	Me.FromString Me.Scheme & "://" & Me.Userinfo & "@" & strHost & ":" & Me.Port & Me.Resource
End Property

Public Property Get Hostname()
	Hostname = pHostname
End Property

Public Property Let Hostname(strHostname)
	Me.FromString Me.Scheme & "://" & Me.Userinfo & "@" & strHostname & ":" & Me.Port & Me.Resource
End Property

Public Property Get Subdomain()
	Subdomain = pSubdomain
End Property

Public Property Let Subdomain(strSubdomain)
	Me.FromString Me.Scheme & "://" & Me.Userinfo & "@" & strSubdomain & "." & Me.Domain & "." & Me.TLD & ":" & Me.Port & Me.Resource
End Property

Public Property Get Domain()
	Domain = pDomain
End Property

Public Property Let Domain(strDomain)
	Me.FromString Me.Scheme & "://" & Me.Userinfo & "@" & Me.Subdomain & "." & strDomain & "." & Me.TLD & ":" & Me.Port & Me.Resource
End Property

Public Property Get TLD()
	TLD = pTLD
End Property

Public Property Let TLD(strTLD)
	Me.FromString Me.Scheme & "://" & Me.Userinfo & "@" & Me.Subdomain & "." & Me.Domain & "." & strTLD & ":" & Me.Port & Me.Resource
End Property

Public Property Get Username()
	Username = pUsername
End Property

Public Property Let Username(strUsername)
	Me.FromString Me.Scheme & "://" & strUsername & ":" & Me.Password & "@" & Me.Host & ":" & Me.Port & Me.Resource
End Property

Public Property Get Password()
	Password = pPassword
End Property

Public Property Let Password(strPassword)
	Me.FromString Me.Scheme & "://" & Me.Username & ":" & strPassword & "@" & Me.Host & ":" & Me.Port & Me.Resource
End Property

Public Property Get Userinfo()
	Userinfo = pUsername & ":" & pPassword
End Property

Public Property Let Userinfo(strUserinfo)
	Me.FromString Me.Scheme & "://" & strUserinfo & "@" & Me.Host & ":" & Me.Port & Me.Resource
End Property 

Public Property Get Authority()
	Authority = pAuthority
End Property

Public Property Let Authority(strAuthority)
	Me.FromString Me.Scheme & "://" & strAuthority & Me.Resource
End Property

Public Property Get Origin()
	Origin = pScheme & "://" & pAuthority
End Property

Public Property Let Origin(strOrigin)
	Me.FromString strOrigin & Me.Resource
End Property

Public Property Get Port()
	Port = pPort
End Property

Public Property Let Port(strPort)
	Me.FromString Me.Scheme & "://" & Me.Userinfo & "@" & Me.Host & ":" & strPort & Me.Resource
End Property

Public Property Get Path()
	Path = pPath
End Property

Public Property Let Path(strPath)
	Me.FromString Me.Scheme & "://" & Me.Authority & strPath & "?" & Me.Query & "#" & Me.Fragment
End Property

Public Property Get Directory()
	Directory = pDirectory
End Property

Public Property Let Directory(strDirectory)
	Me.FromString Me.Scheme & "://" & Me.Authority & strDirectory & Me.Filename & "?" & Me.Query & "#" & Me.Fragment
End Property

Public Property Get Filename()
	Filename = pFilename
End Property

Public Property Let Filename(strFilename)
	Me.FromString Me.Scheme & "://" & Me.Authority & Me.Directory & "/" & strFilename & "?" & Me.Query & "#" & Me.Fragment
End Property

Public Property Get Suffix()
	Suffix = pSuffix
End Property

Public Property Let Suffix(strSuffix)
	If Not InStr(strSuffix, ".") > 0 Then strSuffix = "." & strSuffix
	Me.FromString Me.Scheme & "://" & Me.Authority & Me.Directory & "/" & Split(Me.Filename, ".")(0) & strSuffix & "?" & Me.Query & "#" & Me.Fragment
End Property

Public Property Get Search()
	Search = pQuery
End Property

Public Property Let Search(strSearch)
	If Not InStr(strSearch, "?") > 0 Then strSearch = "?" & strSearch
	Me.FromString Me.Scheme & "://" & Me.Authority & Me.Path & strSearch & "#" & Me.Fragment 
End Property

Public Property Get Query()
	Query = pQuery
End Property

Public Property Let Query(strQuery)
	If Not InStr(strQuery, "?") > 0 Then strQuery = "?" & strQuery
	Me.FromString Me.Scheme & "://" & Me.Authority & Me.Path & strQuery & "#" & Me.Fragment 
End Property

Public Property Get Fragment()
	Fragment = pFragment
End Property

Public Property Let Fragment(strFragment)
	If Not InStr(strFragment, "#") > 0 Then strFragment = "#" & strFragment
	Me.FromString Me.Scheme & "://" & Me.Authority & Me.Path & "?" & Me.Query & strFragment
End Property

Public Property Get Hash()
	Hash = pFragment
End Property

Public Property Let Hash(strHash)
	If Not InStr(strHash, "#") > 0 Then strHash = "#" & strHash
	Me.FromString Me.Scheme & "://" & Me.Authority & Me.Path & "?" & Me.Query & strHash 
End Property

Public Property Get Resource()
	Resource = pPath & _
			IIf(pQuery <> "", "?" & pQuery, pQuery) & _
			IIf(pFragment <> "", "#" & pFragment, pFragment) 
End Property

Public Property Let Resource(strResource)
	Me.FromString Me.Origin & strResource
End Property


' Methods


Public Default Sub Make(strScheme, strAuthority, strPath, strQuery, strFragment)		
	FromString(strScheme & "://" & strAuthority & strPath & _
			IIf(strQuery <> "", "?" & strQuery, strQuery) & _
			IIf(strFragment <> "", "#" & strFragment, strFragment))
End Sub

Public Sub FromString(strURL)
	Class_Initialize()

	If TypeName(strURL) = "String" And strURL <> "" Then
		pUriParser.href = strURL

		pHostname = pUriParser.hostname

		With pUriRegex
			.Pattern = "^(\w+:\/\/)?((.*?)(:(.*?)|)@)"
			If .Test(strURL) Then
				Dim userpassMatches
				Set userpassMatches = .Execute(strURL)
				pUsername = userpassMatches.Item(0).Submatches.Item(2)
				pPassword = userpassMatches.Item(0).Submatches.Item(4)
				Set userpassMatches = Nothing
			End If
		End With

		If pUsername <> "" And pPassword <> "" Then
			pAuthority = pUsername & ":" & pPassword & "@" & pUriParser.host
		Else
			pAuthority = pUriParser.host
		End If

		With pUriRegex
			.Pattern = "(?:([a-z0-9\.\-]*)\.)?((?!com)[a-z0-9\-]{3,}(?=\.[a-z\.]{2,}))\.(?:([a-z\.]{2,})$)"
			If .Test(pHostname) Then
				Dim domainMatches
				Set domainMatches = .Execute(pHostname)
				pSubdomain = domainMatches.Item(0).Submatches.Item(0)
				pDomain = domainMatches.Item(0).Submatches.Item(1)
				pTLD = domainMatches.Item(0).Submatches.Item(2)
				Set domainMatches = Nothing
			End If
		End With

		If pUriParser.protocol <> "" Then pScheme = Left(pUriParser.protocol, Len(pUriParser.protocol) - 1)

		pSchemeName 	= pUriParser.protocolLong
		pPort 		= pUriParser.port
		pPath 		= "/"

		If pUriParser.pathname <> "" Then
			pPath = pPath & pUriParser.pathname

			Dim arrPath
			arrPath = Split(pPath, "/")

			If InStr(arrPath(UBound(arrPath)), ".") > 0 Then
				pFilename = arrPath(UBound(arrPath))
				pSuffix = Split(pFilename, ".")(1)
				ReDim Preserve arrPath(UBound(arrPath) - 1)
			End If

			pDirectory = Join(arrPath, "/")
		End If

		If pUriParser.search <> "" Then pQuery = Right(pUriParser.search, Len(pUriParser.search) - 1)
		If pUriParser.hash <> "" Then pFragment = Right(pUriParser.hash, Len(pUriParser.hash) - 1)
	End If
End Sub

Public Function ToString()
	ToString = pScheme & "://" & pAuthority & pPath & _
			IIf(pQuery <> "", "?" & pQuery, pQuery) & _
			IIf(pFragment <> "", "#" & pFragment, pFragment) 
End Function

Public Function ToArray()
	Dim arrProperties, _
		arrArray(), _
		intCount, _
		i

	arrProperties = Array(Me.Scheme, _
				Me.Username, _
				Me.Password, _
				Me.Subdomain, _
				Me.Domain, _
				Me.TLD, _
				Me.Port, _
				Me.Directory, _
				Me.Filename, _
				Me.Query, _
				Me.Fragment)

	intCount = 0

	For i = 0 To UBound(arrProperties)
		If arrProperties(i) <> "" Then intCount = intCount + 1
	Next

	If intCount > 0 Then
		ReDim arrArray(intCount - 1)

		intCount = 0

		For i = 0 To UBound(arrProperties)
			If arrProperties(i) <> "" Then
				intCount = intCount + 1
				arrArray(intCount - 1) = arrProperties(i)
			End If
		Next

		ToArray = arrArray
	Else
		ToArray = Array()
	End If
End Function

Public Function ToDict()
	Dim objDict
	Set objDict = CreateObject("Scripting.Dictionary")

	objDict.Add "Scheme", Me.Scheme
	objDict.Add "SchemeName", Me.SchemeName
	objDict.Add "Protocol", Me.Protocol
	objDict.Add "ProtocolName", Me.ProtocolName
	objDict.Add "Host", Me.Host
	objDict.Add "Hostname", Me.Hostname
	objDict.Add "Subdomain", Me.Subdomain
	objDict.Add "Domain", Me.Domain
	objDict.Add "TLD", Me.TLD
	objDict.Add "Username", Me.Username
	objDict.Add "Password", Me.Password
	objDict.Add "Userinfo", Me.Userinfo
	objDict.Add "Authority", Me.Authority
	objDict.Add "Origin", Me.Origin
	objDict.Add "Port", Me.Port
	objDict.Add "Path", Me.Path
	objDict.Add "Directory", Me.Directory
	objDict.Add "Filename", Me.Filename
	objDict.Add "Suffix", Me.Suffix
	objDict.Add "Search", Me.Search
	objDict.Add "Query", Me.Query
	objDict.Add "Fragment", Me.Fragment
	objDict.Add "Hash", Me.Hash
	objDict.Add "Resource", Me.Resource

	Set ToDict = objDict
End Function

Public Function Segment()
	Dim arrSegment

	If Left(pPath, 1) = "/" Then
		arrSegment = Split(Mid(pUriCoder.Run("decode", Array(pPath)), 2, Len(pPath)), "/")
	Else
		arrSegment = Split(pUriCoder.Run("decode", Array(pPath)), "/")
	End If

	Segment = arrSegment
End Function

Public Function SegmentCoded()
	Dim arrSegment, _
		i

	arrSegment = Me.Segment()

	For i = 0 To UBound(arrSegment)
		arrSegment(i) = pUriCoder.Run("encode", Array(arrSegment(i)))
	Next

	SegmentCoded = arrSegment
End Function

Public Sub Replace(strSegment, strReplace)
	ModifySegment strSegment, strReplace, "", True, False
End Sub

Public Sub Add(strSegment, strValue)
	ModifySegment strSegment, strValue, "", False, False
End Sub

Public Sub AddTo(strSegment, strAppend)
	ModifySegment strSegment, strAppend, "", False, False
End Sub

Public Sub Remove(strSegment)
	ModifySegment strSegment, "", "", True, False
End Sub

Public Sub RemoveFrom(strSegment, strRemove)
	ModifySegment strSegment, "", strRemove, True, True
End Sub

Public Function Encode()
	Encode = pUriCoder.Run("encode", Array(Me.ToString()))
End Function

Public Function Decode()
	Decode = pUriCoder.Run("decode", Array(Me.ToString()))
End Function

Public Function Validate()
	With pUriRegex
		.Pattern = "^(?:(?:https?|ftp):\/\/)(?:\S+(?::\S*)?@)?(?:(?!(?:10|127)" & _
				"(?:\.\d{1,3}){3})(?!(?:169\.254|192\.168)(?:\.\d{1,3}){2})" & _
				"(?!172\.(?:1[6-9]|2\d|3[0-1])(?:\.\d{1,3}){2})" & _
				"(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5]))" & _
				"{2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]-*)" & _
				"*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)" & _
				"*(?:\.(?:[a-z\u00a1-\uffff]{2,}))\.?)(?::\d{2,5})?(?:[/?#]\S*)?$"
		.IgnoreCase = True
		Validate = .Test(Me.ToString())
	End With
End Function


' Helper Methods


Private Sub ModifySegment(strSegment, strValue, strFind, blnReplace, blnExists)
	Select Case LCase(strSegment)
		Case "scheme":
			If blnExists And Me.Scheme = "" Then Exit Sub
			Me.Scheme = SegmentNewValue(Me.Scheme, strValue, strFind, blnReplace)
		Case "schemename":
			If blnExists And Me.SchemeName = "" Then Exit Sub
			Me.SchemeName = SegmentNewValue(Me.SchemeName, strValue, strFind, blnReplace)
		Case "protocol":
			If blnExists And Me.Protocol = "" Then Exit Sub
			Me.Protocol = SegmentNewValue(Me.Protocol, strValue, strFind, blnReplace)
		Case "protocolname":
			If blnExists And Me.ProtocolName = "" Then Exit Sub
			Me.ProtocolName = SegmentNewValue(Me.ProtocolName, strValue, strFind, blnReplace)
		Case "host":
			If blnExists And Me.Host = "" Then Exit Sub
			Me.Host = SegmentNewValue(Me.Host, strValue, strFind, blnReplace)
		Case "hostname":
			If blnExists And Me.Hostname = "" Then Exit Sub
			Me.Hostname = SegmentNewValue(Me.Hostname, strValue, strFind, blnReplace)
		Case "subdomain":
			If blnExists And Me.Subdomain = "" Then Exit Sub
			Me.Subdomain = SegmentNewValue(Me.Subdomain, strValue, strFind, blnReplace)
		Case "domain":
			If blnExists And Me.Domain = "" Then Exit Sub
			Me.Domain = SegmentNewValue(Me.Domain, strValue, strFind, blnReplace)
		Case "tld":
			If blnExists And Me.TLD = "" Then Exit Sub
			Me.TLD = SegmentNewValue(Me.TLD, strValue, strFind, blnReplace)
		Case "username":
			If blnExists And Me.Username = "" Then Exit Sub
			Me.Username = SegmentNewValue(Me.Username, strValue, strFind, blnReplace)
		Case "password":
			If blnExists And Me.Password = "" Then Exit Sub
			Me.Password = SegmentNewValue(Me.Password, strValue, strFind, blnReplace)
		Case "userinfo":
			If blnExists And Me.Userinfo = "" Then Exit Sub
			Me.Userinfo = SegmentNewValue(Me.Userinfo, strValue, strFind, blnReplace)
		Case "authority":
			If blnExists And Me.Authority = "" Then Exit Sub
			Me.Authority = SegmentNewValue(Me.Authority, strValue, strFind, blnReplace)
		Case "origin":
			If blnExists And Me.Origin = "" Then Exit Sub
			Me.Origin = SegmentNewValue(Me.Origin, strValue, strFind, blnReplace)
		Case "port":
			If blnExists And Me.Port = "" Then Exit Sub
			Me.Port = SegmentNewValue(Me.Port, strValue, strFind, blnReplace)
		Case "path":
			If blnExists And Me.Path = "" Then Exit Sub
			Me.Path = SegmentNewValue(Me.Path, strValue, strFind, blnReplace)
		Case "directory":
			If blnExists And Me.Directory = "" Then Exit Sub
			Me.Directory = SegmentNewValue(Me.Directory, strValue, strFind, blnReplace)
		Case "filename":
			If blnExists And Me.Filename = "" Then Exit Sub
			Me.Filename = SegmentNewValue(Me.Filename, strValue, strFind, blnReplace)
		Case "suffix":
			If blnExists And Me.Suffix = "" Then Exit Sub
			Me.Suffix = SegmentNewValue(Me.Suffix, strValue, strFind, blnReplace)
		Case "search":
			If blnExists And Me.Search = "" Then Exit Sub
			If Left(strValue, 1) = "?" Then strValue = Right(strValue, Len(strValue) - 1)
			Me.Search = SegmentNewValue(Me.Search, strValue, strFind, blnReplace)
		Case "query":
			If blnExists And Me.Query = "" Then Exit Sub
			If Left(strValue, 1) = "?" Then strValue = Right(strValue, Len(strValue) - 1)
			Me.Query = SegmentNewValue(Me.Query, strValue, strFind, blnReplace)
		Case "fragment":
			If blnExists And Me.Fragment = "" Then Exit Sub
			Me.Fragment = SegmentNewValue(Me.Fragment, strValue, strFind, blnReplace)
		Case "hash":
			If blnExists And Me.Hash = "" Then Exit Sub
			Me.Hash = SegmentNewValue(Me.Hash, strValue, strFind, blnReplace)
		Case "resource":
			If blnExists And Me.Resource = "" Then Exit Sub
			Me.Resource = SegmentNewValue(Me.Resource, strValue, strFind, blnReplace)
	End Select
End Sub

Private Function SegmentNewValue(strSegment, strValue, strFind, blnReplace)
	If strFind <> "" Then
		Dim intPosition

		intPosition = InStr(strSegment, strFind)

		If blnReplace Then
			strSegment = Left(strSegment, intPosition - 1) & strValue & Right(strSegment, Len(strSegment) - intPosition - Len(strFind) + 1)
		Else
			strSegment = Left(strSegment, intPosition + Len(strFind)) & strValue & Right(strSegment, Len(strSegment) - intPosition - Len(strFind) + 1)
		End If
	Else
		If blnReplace Then
			strSegment = strValue
		Else
			strSegment = strSegment & strValue
		End If
	End If

	SegmentNewValue = strSegment
End Function

Private Sub Class_Terminate()
	Set pUriRegex = Nothing
	Set pUriParserHost = Nothing
	Set pUriParser = Nothing
End Sub
