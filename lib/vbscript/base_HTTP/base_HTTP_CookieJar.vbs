Option Explicit

Include "base_HTTP_Cookie"

Class base_HTTP_CookieJar
	Private p_objCookies, _
		p_strFilename

	Private Sub Class_Initialize()
		Set p_objCookies = CreateObject("Scripting.Dictionary")
	End Sub


	' Properties


	Public Default Property Get Cookie(name)
		Set Cookie = p_objCookies(name)
	End Property

	Public Property Get Cookies()

	End Property

	Public Property Get Count()
		Count = p_objCookies.Count
	End Property

	Public Property Get Filename()

	End Property

	Public Sub Add( _
		varHttpCookie _
		)

		Dim objCookie

		If TypeName(varHttpCookie) = "base_HTTP_Cookie" Then
			Set objCookie = varHttpCookie
		Else
			Set objCookie = New base_HTTP_Cookie

			If TypeName(varHttpCookie) = "String" Then
				objCookie.FromString(varHttpCookie)
			ElseIf TypeName(varHttpCookie) = "Dictionary" Then
				objCookie.FromDict(varHttpCookie)
			Else
				' Error
			End If
		End If

		p_objCookies.Add objCookie.Name, objCookie
	End Sub

	Public Sub Remove( _
    		ByVal strKey _
		)

	End Sub

	Public Function ExtractCookies()

	End Function

	Public Sub Load( _
		ByVal strFilename _
		)

	End Sub

	Public Sub Save()

	End Sub

	Public Sub SaveAs( _
		ByVal strFilename _
		)

	End Sub

	Public Sub Revert()

	End Sub

	Public Sub Clear()

	End Sub

	Public Sub ClearTemporary()

	End Sub

	Public Sub Match()

	End Sub

	Public Function ToString()
		Dim varKey, _
			strCookie
    
		strCookie = ""
    
		If p_objCookieDict.Count > 0 Then
			For Each varKey In p_objCookieDict.Keys()
				strCookie = strCookie & varKey & "=" & p_objCookieDict(varKey) & "; "
			Next
        
			strCookie = Left(strCookie, Len(strCookie) - 2)
		End If
    
		ToString = strCookie
	End Function

	Public Function FromResponseHeaders( _
    		ByVal varResponseHeaders _
		)
    
		Dim arrCookie As Variant, _
			strCookie As String, _
			i
    
		arrCookie = Split(strCookies, vbCrLf)
    
		For i = 0 To UBound(arrCookie) - 2
			If Split(arrCookie(i), ": ")(0) = "Set-Cookie" Then
				strCookie = Split(Split(arrCookie(i), ": ")(1), "; ")(0)
				Me.Add strCookie
			End If
    		Next
	End Function

	Private Sub Class_Terminate()
		Set p_objCookies = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_HTTP_CookieJar.vbs" Then

End If
