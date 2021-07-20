Option Explicit

Include "base_Data_Dictionary"
Include "base_Sys_Script"
Include "base_URI"

Class base_HTTP_Cookie
	Private p_strName, _
		p_strValue, _
		p_strExpires, _
		p_strDomain, _
		p_strPath, _
		p_blnSecure, _
		p_blnHttpOnly, _
		p_varPort, _
		p_strComment, _
		p_strCommentUrl, _
		p_blnDiscard, _
		p_intVersion

	Private Sub Class_Initialize()
		p_strName = ""
		p_strValue = ""
		p_strExpires = ""
		p_strDomain = ""
		p_strPath = ""
		p_blnSecure = False
		p_blnHttpOnly = False
		p_varPort = 0
		p_strComment = ""
		p_strCommentUrl = ""
		p_blnDiscard = False
		p_intVersion = 1
	End Sub


	' Constructor


	Public Default Function Make( _
    		ByVal strName, _
    		ByVal strValue, _
    		ByVal dtmExpires, _
    		ByVal strDomain, _
    		ByVal strPath, _
    		ByVal blnSecure, _
    		ByVal blnHttpOnly _
    		)

		p_strName = strName
		p_strValue = strValue
		p_strExpires = dtmExpires
		p_strDomain = strDomain
		p_strPath = strPath
		p_blnSecure = blnSecure
		p_blnHttpOnly = blnHttpOnly
		Set Make = Me		
	End Function


	' Properties


	Public Property Get Attributes()
		Dim objCookieDict

		Set objCookieDict = New base_Data_Dictionary

		With objCookieDict
			.Add("Name") = p_strName
			.Add("Value") = p_strValue
			.Add("Expires") = p_strExpires
			.Add("Domain") = p_strValue
			.Add("Path") = p_strPath
			.Add("Secure") = p_blnSecure
			.Add("HttpOnly") = p_blnHttpOnly
			.Add("Port") = p_varPort
			.Add("Comment") = p_strComment
			.Add("CommentUrl") = p_strCommentUrl
			.Add("Discard") = p_blnDiscard
			.Add("Version") = p_intVersion
		End With

		Set Attributes = objCookieDict
	End Property

	Public Property Get Name()
		Name = p_strName
	End Property

	Public Property Get Value()
		Value = p_strValue
	End Property

	Public Property Get Port()
		Port = p_varPort
	End Property

	Public Property Get Domain()
		Domain = p_strDomain
	End Property

	Public Property Get Path()
		Path = p_strPath
	End Property

	Public Property Get Expires()
		Expires = p_strExpires
	End Property

	Public Property Get Discard()
		Discard = p_blnDiscard
	End Property

	Public Property Get Comment()
		Comment = p_strComment
	End Property

	Public Property Get CommentUrl()
		CommentUrl = p_strCommentUrl
	End Property

	Public Property Get Secure()
		Secure = p_blnSecure
	End Property

	Public Property Get HttpOnly()
		HttpOnly = p_blnHttpOnly
	End Property

	Public Property Get Version()
		Version = p_intVersion
	End Property

	Public Property Get IsPortSpecified()
		If Not IsNull(p_varPort) And p_varPort <> 0 Then
			IsPortSpecified = True
		Else
			IsPortSpecified = False
    		End If
	End Property

	Public Property Get IsDomainSpecified()
		If Not IsNull(p_strDomain) And TypeName(p_strDomain) = "String" Then
			IsDomainSpecified = True
		Else
			IsDomainSpecified = False
		End If
	End Property

	Public Property Get IsExpired()
		If p_strExpires <> "" Then

			Dim objScript, _
				dtmExpires

			Set objScript = New base_Sys_Script
						
			With objScript
				.Language = "JScript"
				.AddCode("var dt = new Date(""" & p_strExpires & """).toLocaleString();")
			End With

			dtmExpires = CDate(objScript.Variable("dt"))

			If Now() > dtmExpires Then
		 		IsExpired = True
			Else
				IsExpired = False
			End If
		Else
			IsExpired = False
		End If

		Set objScript = Nothing
	End Property


	' Methods


	Public Function Match( _
		ByVal strUrl _
		)

		If Me.IsExpired Then
			Match = False
			Exit Function
		End If

		Dim objUrl
		Set objUrl = New base_URI

		objUrl.FromString strUrl

		If Me.Secure And objUrl.Protocol <> "https" Then
			Match = False
			Exit Function
		End If
	
		If DomainMatch(objUrl.Hostname) And PathMatch(objUrl.Path) Then
			Match = True
		Else
			Match = False
		End If

		Set objUrl = Nothing
	End Function

	Public Function ToString()
		Dim strCookie

		strCookie = p_strName & "=" & p_strValue

		If p_strExpires <> "" Then strCookie = strCookie & "; Expires=" & p_strExpires
		If p_strDomain <> "" Then strCookie = strCookie & "; Domain=" & p_strDomain
		If p_strPath <> "" Then strCookie = strCookie & "; Path=" & p_strPath
		If p_blnSecure <> False Then strCookie = strCookie & "; Secure"
		If p_blnHttpOnly <> False Then strCookie = strCookie & "; HttpOnly"

		' p_varPort
		' p_strComment
		' p_strCommentUrl
		' p_blnDiscard
		' p_intVersion

		ToString = strCookie
	End Function

	Public Sub FromString( _
		ByVal strCookie _
		)

		Class_Initialize()

		Dim arrCookie, _
			intIndex

		arrCookie = Split(strCookie, ";")

		p_strName = Split(arrCookie(0), "=")(0)
		p_strValue = Split(arrCookie(0), "=")(1)

		For intIndex = 1 To UBound(arrCookie)
			If InStr(arrCookie(intIndex), "=") > 0 Then
				Dim arrCookieAttr

				arrCookieAttr = Split(arrCookie(intIndex), "=")

				Select Case LCase(Trim(arrCookieAttr(0)))
					Case "domain":
						' For security reasons, cookies can only be set on the current resource's top
						' domain and its subdomains, and not for another domain and its subdomains.
						' The cookie domain should not have a leading dot, as in .foo.com - simply use foo.com
						p_strDomain = CStr(arrCookieAttr(1))
					Case "path":
						p_strPath = CStr(arrCookieAttr(1))
					Case "expires":
						p_strExpires = CStr(arrCookieAttr(1))
					Case "max-age":
						' p_strExpires = DateAdd("s", arrCookieAttr(1), Now())
					Case "version":
						p_intVersion = CInt(arrCookieAttr(1))
					Case "port":
						p_varPort = arrCookieAttr(1)
					Case "commment":
						p_strComment = arrCookieAttr(1)
					Case "commenturl":
						p_strCommentUrl = arrCookieAttr(1)
				End Select
			Else
				Select Case LCase(Trim(arrCookie(intIndex)))
					Case "httponly":
						p_blnHttpOnly = True
					Case "secure":
						p_blnSecure = True
					Case "discard":
						p_blnDiscard = True
				End Select
			End If
		Next
	End Sub

	Public Function ToDict()

	End Function

	Public Sub FromDict( _
		ByVal objCookieDict _
		)

	End Sub

	Public Function ToArray()

	End Function

	Public Function FromArray()

	End Function


	' Helper Methods


	Private Function DomainMatch( _
		ByVal strDomain _
		)

		Dim blnDomainMatch, _
			arrUrlDomainParts, _
			arrCookieDomainParts, _
			intDomainPartDifferential, _
			intDomainIndex

		blnDomainMatch = False

		arrUrlDomainParts = Split(strDomain, ".")

		If Left(Me.Domain, 1) = "." Then
			arrCookieDomainParts = Split(Right(Me.Domain, Len(Me.Domain) - 1), ".")
		Else
			arrCookieDomainParts = Split(Me.Domain, ".")
		End If

		intDomainPartDifferential = UBound(arrUrlDomainParts) - UBound(arrCookieDomainParts)

		For intDomainIndex = UBound(arrCookieDomainParts) To 0 Step -1
			If arrCookieDomainParts(intDomainIndex) = arrUrlDomainParts(intDomainIndex + intDomainPartDifferential) Then
				blnDomainMatch = True
			Else
				blnDomainMatch = False
				Exit For
			End If
		Next

		DomainMatch = blnDomainMatch
	End Function

	Private Function PathMatch( _
		ByVal strPath _
		)

		Dim blnPathMatch

		blnPathMatch = False

		If Me.Path = "" Or Me.Path = "/" Or Me.Path = strPath Then
			blnPathMatch = True
		Else
			Dim arrUrlPathParts, _
				arrCookiePathParts, _
				intPathIndex

			arrUrlPathParts = Split(strPath, "/")
			arrCookiePathParts = Split(Me.Path, "/")

			For intPathIndex = 1 To UBound(arrCookiePathParts)
				If intPathIndex > UBound(arrUrlPathParts) Then 
					blnPathMatch = False
					Exit For						
				End If

				If arrCookiePathParts(intPathIndex) = arrUrlPathParts(intPathIndex) Then
					blnPathMatch = True
				Else
					blnPathMatch = False
					Exit For
				End If
			Next
		End If

		PathMatch = blnPathMatch
	End Function

	Private Sub Class_Terminate()

	End Sub
End Class

If WScript.ScriptName = "base_HTTP_Cookie.vbs" Then

End If
