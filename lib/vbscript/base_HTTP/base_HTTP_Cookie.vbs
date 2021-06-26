Option Explicit

Include "base_URI"

Class base_HTTP_Cookie
	Private p_objDomainRegex

	Private p_strName, _
		p_strValue, _
		p_intPort, _
		p_strDomain, _
		p_strPath, _
		p_dtmExpires, _
		p_strComment, _
		p_strCommentUrl, _
		p_blnSecure, _
		p_blnHttpOnly, _
		p_blnDiscard, _
		p_intVersion

	Private Sub Class_Initialize()
		Set p_objDomainRegex = New RegExp

		p_strName = ""
		p_strValue = ""
		p_intPort = 0
		p_strDomain = ""
		p_strPath = ""
		p_dtmExpires = 0
		p_strComment = ""
		p_strCommentUrl = ""
		p_blnSecure = False
		p_blnHttpOnly = False
		p_blnDiscard = False
		p_intVersion = 1
	End Sub


	' Constructor


	Public Default Function Make( _
    		ByVal strName, _
    		ByVal strValue, _
    		ByVal strDomain, _
    		ByVal strPath, _
    		ByVal dtmExpires, _
    		ByVal strComment, _
    		ByVal blnSecure, _
    		ByVal blnHttpOnly, _
    		ByVal intVersion _
    		)

		p_strName = strName
		p_strValue = strValue
		p_strDomain = strDomain
		p_strPath = strPath
		p_dtmExpires = dtmExpires
		p_strComment = strComment
		p_blnSecure = blnSecure
		p_blnHttpOnly = blnHttpOnly
		p_intVersion = intVersion
		Set Make = Me		
	End Function


	' Properties


	Public Property Get Attributes()

	End Property

	Public Property Get Name()
		Name = p_strName
	End Property

	Public Property Get Value()
		Value = p_strValue
	End Property

	Public Property Get Port()
		Port = p_intPort
	End Property

	Public Property Get Domain()
		Domain = p_strDomain
	End Property

	Public Property Get Path()
		Path = p_strPath
	End Property

	Public Property Get Expires()
		Expires = p_dtmExpires
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
		If Not IsNull(p_intPort) And IsNumeric(p_intPort) Then
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
		If Now() > p_dtmExpires Then
			IsExpired = True
		Else
			IsExpired = False
		End If
	End Property


	' Methods


	Public Function Match(strURL)

	End Function

	Public Function ToString()

	End Function

	Public Function FromString( _
		ByVal strCookie _
		)

	End Function

	Public Function ToDict()

	End Function

	Public Function FromDict( _
		ByVal objCookieDict _
		)

	End Function

	Public Function ToArray()

	End Function

	Public Function FromArray()

	End Function

	Private Sub Class_Terminate()
		Set p_objDomainRegex = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_HTTP_Cookie.vbs" Then

End If
