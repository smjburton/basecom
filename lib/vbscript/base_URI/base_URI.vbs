Option Explicit

Include "base_Sys_Util"
Include "base_Sys_Script"

Class base_URI
	Private p_objUriRegex, _
		p_objUriParserHost, _
		p_objUriParser, _
		p_objUriCoder

	Private p_strScheme, _
		p_strSchemeName, _
		p_strHostname, _
		p_strSubdomain, _
		p_strDomain, _
		p_strTLD, _
		p_strUsername, _
		p_strPassword, _
		p_strUserinfo, _
		p_strAuthority, _
		p_intPort, _
		p_strPath, _
		p_strDirectory, _
		p_strFilename, _
		p_strSuffix, _
		p_strQuery, _
		p_strFragment

	Private Sub Class_Initialize()
		Set p_objUriRegex = New RegExp
		Set p_objUriParserHost = CreateObject("HTMLFile")
		Set p_objUriParser = p_objUriParserHost.createElement("a")
		Set p_objUriCoder = New base_Sys_Script

		With p_objUriCoder
			.Language = "JScript"
			.AddCode "function encode(uri) { return encodeURIComponent(uri); }"
			.AddCode "function decode(uri) { return decodeURIComponent(uri); }"
		End With

		p_strScheme = ""
		p_strSchemeName = ""
		p_strHostname = ""
		p_strSubdomain = ""
		p_strDomain = ""
		p_strTLD = ""
		p_strUsername = ""
		p_strPassword = ""
		p_strUserinfo = ""
		p_strAuthority = ""
		p_intPort = ""
		p_strPath = ""
		p_strDirectory = ""
		p_strFilename = ""
		p_strSuffix = ""
		p_strQuery = ""
		p_strFragment = ""
	End Sub


	' Properties


	Public Property Get Scheme()
		Scheme = p_strScheme
	End Property

	Public Property Let Scheme(strScheme)
		Me.FromString strScheme & "://" & Me.Authority & Me.Resource
	End Property

	Public Property Get SchemeName()
		SchemeName = p_strSchemeName
	End Property	

	Public Property Get Protocol()
		Protocol = p_strScheme
	End Property

	Public Property Let Protocol(strProtocol)
		Me.FromString strProtocol & "://" & Me.Authority & Me.Resource
	End Property

	Public Property Get ProtocolName()
		ProtocolName = p_strSchemeName
	End Property

	Public Property Get Host()
		Host = p_strHostname
	End Property

	Public Property Let Host(strHost)
		Me.FromString Me.Scheme & "://" & Me.Userinfo & "@" & strHost & ":" & Me.Port & Me.Resource
	End Property

	Public Property Get Hostname()
		Hostname = p_strHostname
	End Property

	Public Property Let Hostname(strHostname)
		Me.FromString Me.Scheme & "://" & Me.Userinfo & "@" & strHostname & ":" & Me.Port & Me.Resource
	End Property

	Public Property Get Subdomain()
		Subdomain = p_strSubdomain
	End Property

	Public Property Let Subdomain(strSubdomain)
		Me.FromString Me.Scheme & "://" & Me.Userinfo & "@" & strSubdomain & "." & Me.Domain & "." & Me.TLD & ":" & Me.Port & Me.Resource
	End Property

	Public Property Get Domain()
		Domain = p_strDomain
	End Property

	Public Property Let Domain(strDomain)
		Me.FromString Me.Scheme & "://" & Me.Userinfo & "@" & Me.Subdomain & "." & strDomain & "." & Me.TLD & ":" & Me.Port & Me.Resource
	End Property

	Public Property Get TLD()
		TLD = p_strTLD
	End Property

	Public Property Let TLD(strTLD)
		Me.FromString Me.Scheme & "://" & Me.Userinfo & "@" & Me.Subdomain & "." & Me.Domain & "." & strTLD & ":" & Me.Port & Me.Resource
	End Property

	Public Property Get Username()
		Username = p_strUsername
	End Property

	Public Property Let Username(strUsername)
		Me.FromString Me.Scheme & "://" & strUsername & ":" & Me.Password & "@" & Me.Host & ":" & Me.Port & Me.Resource
	End Property

	Public Property Get Password()
		Password = p_strPassword
	End Property

	Public Property Let Password(strPassword)
		Me.FromString Me.Scheme & "://" & Me.Username & ":" & strPassword & "@" & Me.Host & ":" & Me.Port & Me.Resource
	End Property

	Public Property Get Userinfo()
		Userinfo = p_strUsername & ":" & p_strPassword
	End Property

	Public Property Let Userinfo(strUserinfo)
		Me.FromString Me.Scheme & "://" & strUserinfo & "@" & Me.Host & ":" & Me.Port & Me.Resource
	End Property 

	Public Property Get Authority()
		Authority = p_strAuthority
	End Property

	Public Property Let Authority(strAuthority)
		Me.FromString Me.Scheme & "://" & strAuthority & Me.Resource
	End Property

	Public Property Get Origin()
		Origin = p_strScheme & "://" & p_strAuthority
	End Property

	Public Property Let Origin(strOrigin)
		Me.FromString strOrigin & Me.Resource
	End Property

	Public Property Get Port()
		Port = p_intPort
	End Property

	Public Property Let Port(strPort)
		Me.FromString Me.Scheme & "://" & Me.Userinfo & "@" & Me.Host & ":" & strPort & Me.Resource
	End Property

	Public Property Get Path()
		Path = p_strPath
	End Property

	Public Property Let Path(strPath)
		Me.FromString Me.Scheme & "://" & Me.Authority & strPath & "?" & Me.Query & "#" & Me.Fragment
	End Property

	Public Property Get Directory()
		Directory = p_strDirectory
	End Property

	Public Property Let Directory(strDirectory)
		Me.FromString Me.Scheme & "://" & Me.Authority & strDirectory & Me.Filename & "?" & Me.Query & "#" & Me.Fragment
	End Property

	Public Property Get Filename()
		Filename = p_strFilename
	End Property

	Public Property Let Filename(strFilename)
		Me.FromString Me.Scheme & "://" & Me.Authority & Me.Directory & "/" & strFilename & "?" & Me.Query & "#" & Me.Fragment
	End Property

	Public Property Get Suffix()
		Suffix = p_strSuffix
	End Property

	Public Property Let Suffix(strSuffix)
		If Not InStr(strSuffix, ".") > 0 Then strSuffix = "." & strSuffix
		Me.FromString Me.Scheme & "://" & Me.Authority & Me.Directory & "/" & Split(Me.Filename, ".")(0) & strSuffix & "?" & Me.Query & "#" & Me.Fragment
	End Property

	Public Property Get Search()
		Search = p_strQuery
	End Property

	Public Property Let Search(strSearch)
		If Not InStr(strSearch, "?") > 0 Then strSearch = "?" & strSearch
		Me.FromString Me.Scheme & "://" & Me.Authority & Me.Path & strSearch & "#" & Me.Fragment 
	End Property

	Public Property Get Query()
		Query = p_strQuery
	End Property

	Public Property Let Query(strQuery)
		If Not InStr(strQuery, "?") > 0 Then strQuery = "?" & strQuery
		Me.FromString Me.Scheme & "://" & Me.Authority & Me.Path & strQuery & "#" & Me.Fragment 
	End Property

	Public Property Get Fragment()
		Fragment = p_strFragment
	End Property

	Public Property Let Fragment(strFragment)
		If Not InStr(strFragment, "#") > 0 Then strFragment = "#" & strFragment
		Me.FromString Me.Scheme & "://" & Me.Authority & Me.Path & "?" & Me.Query & strFragment
	End Property

	Public Property Get Hash()
		Hash = p_strFragment
	End Property

	Public Property Let Hash(strHash)
		If Not InStr(strHash, "#") > 0 Then strHash = "#" & strHash
		Me.FromString Me.Scheme & "://" & Me.Authority & Me.Path & "?" & Me.Query & strHash 
	End Property

	Public Property Get Resource()
		Resource = p_strPath & _
				IIf(p_strQuery <> "", "?" & p_strQuery, p_strQuery) & _
				IIf(p_strFragment <> "", "#" & p_strFragment, p_strFragment) 
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
			p_objUriParser.href = strURL

			p_strHostname = p_objUriParser.hostname

			With p_objUriRegex
				.Pattern = "^(\w+:\/\/)?((.*?)(:(.*?)|)@)"

				If .Test(strURL) Then
					Dim userpassMatches
					Set userpassMatches = .Execute(strURL)
					p_strUsername = userpassMatches.Item(0).Submatches.Item(2)
					p_strPassword = userpassMatches.Item(0).Submatches.Item(4)
					Set userpassMatches = Nothing
				End If
			End With

			If p_strUsername <> "" And p_strPassword <> "" Then
				p_strAuthority = p_strUsername & ":" & p_strPassword & "@" & p_objUriParser.host
			Else
				p_strAuthority = p_objUriParser.host
			End If

			With p_objUriRegex
				.Pattern = "(?:([a-z0-9\.\-]*)\.)?((?!com)[a-z0-9\-]{3,}(?=\.[a-z\.]{2,}))\.(?:([a-z\.]{2,})$)"

				If .Test(p_strHostname) Then
					Dim domainMatches
					Set domainMatches = .Execute(p_strHostname)
					p_strSubdomain = domainMatches.Item(0).Submatches.Item(0)
					p_strDomain = domainMatches.Item(0).Submatches.Item(1)
					p_strTLD = domainMatches.Item(0).Submatches.Item(2)
					Set domainMatches = Nothing
				End If
			End With

			If p_objUriParser.protocol <> "" Then p_strScheme = Left(p_objUriParser.protocol, Len(p_objUriParser.protocol) - 1)

			p_strSchemeName = p_objUriParser.protocolLong
			p_intPort = p_objUriParser.port
			p_strPath = "/"

			If p_objUriParser.pathname <> "" Then
				p_strPath = p_strPath & p_objUriParser.pathname

				Dim arrPath
				arrPath = Split(p_strPath, "/")

				If InStr(arrPath(UBound(arrPath)), ".") > 0 Then
					p_strFilename = arrPath(UBound(arrPath))
					p_strSuffix = Split(p_strFilename, ".")(1)
					ReDim Preserve arrPath(UBound(arrPath) - 1)
				End If

				p_strDirectory = Join(arrPath, "/")
			End If

			If p_objUriParser.search <> "" Then p_strQuery = Right(p_objUriParser.search, Len(p_objUriParser.search) - 1)
			If p_objUriParser.hash <> "" Then p_strFragment = Right(p_objUriParser.hash, Len(p_objUriParser.hash) - 1)
		End If
	End Sub

	Public Function ToString()
		ToString = IIf(p_strScheme <> "", p_strScheme & "://", p_strScheme) & p_strAuthority & p_strPath & _
				IIf(p_strQuery <> "", "?" & p_strQuery, p_strQuery) & _
				IIf(p_strFragment <> "", "#" & p_strFragment, p_strFragment) 
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

		If Left(p_strPath, 1) = "/" Then
			arrSegment = Split(Mid(p_objUriCoder.Run("decode", Array(p_strPath)), 2, Len(p_strPath)), "/")
		Else
			arrSegment = Split(p_objUriCoder.Run("decode", Array(p_strPath)), "/")
		End If

		Segment = arrSegment
	End Function

	Public Function SegmentCoded()
		Dim arrSegment, _
			i

		arrSegment = Me.Segment()

		For i = 0 To UBound(arrSegment)
			arrSegment(i) = p_objUriCoder.Run("encode", Array(arrSegment(i)))
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
		Encode = p_objUriCoder.Run("encode", Array(Me.ToString()))
	End Function

	Public Function Decode()
		Decode = p_objUriCoder.Run("decode", Array(Me.ToString()))
	End Function

	Public Function Validate()
		With p_objUriRegex
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
		Set p_objUriRegex = Nothing
		Set p_objUriParserHost = Nothing
		Set p_objUriParser = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_URI.vbs" Then
	Dim URL
	Set URL = New base_URI

	URL.FromString("https://user:pass1234@www.sub.sub.example.co.uk/path/to thing/file.html?search=something&page=1#Header1")

	WScript.Echo URL.Host
End If
