Option Explicit

Include "base_Data_Array"
Include "base_Data_Dictionary"
Include "base_HTTP_Cookie"
Include "base_IO_TextStream"

Private Const FOR_READING = 1
Private Const FOR_WRITING = 2
Private Const FOR_APPENDING = 8

Class base_HTTP_CookieJar
	Private p_objCookies, _
		p_objTextStream, _
		p_strFilename

	Private Sub Class_Initialize()
		Set p_objCookies = New base_Data_Dictionary
		Set p_objTextStream = New base_IO_TextStream
	End Sub


	' Properties


	Public Default Property Get Cookie( _
		ByVal strName _
		)

		Set Cookie = p_objCookies(strName)
	End Property

	Public Property Get Cookies()
		Cookies = p_objCookies.Items()
	End Property

	Public Property Get Count()
		Count = p_objCookies.Count
	End Property

	Public Property Get Filename()
		Filename = p_strFilename
	End Property

	Public Property Let Filename( _
		ByVal strFilename _
		)

		p_strFilename = strFilename
	End Property


	' Methods


	Public Sub Add( _
		ByVal varHttpCookie _
		)
		
		Dim objCookie

		If TypeName(varHttpCookie) = "base_HTTP_Cookie" Then
			Set objCookie = varHttpCookie
		Else
			Set objCookie = New base_HTTP_Cookie

			If TypeName(varHttpCookie) = "String" Then
				objCookie.FromString varHttpCookie
			ElseIf TypeName(varHttpCookie) = "Dictionary" Then
				objCookie.FromDict varHttpCookie
			ElseIf IsArray(varHttpCookie) Then
				objCookie.FromArray varHttpCookie
			Else
				Sys.ErrorHandler.Raise 5000, "base_HTTP_CookieJar.Add()", "Failed to parse cookie from unrecognized variable type: " & TypeName(varHttpCookie) & ".", "", ""
				Exit Sub
			End If
		End If

		If p_objCookies.Exists(objCookie.Name) Then
			Set p_objCookies(objCookie.Name) = objCookie
		Else
			p_objCookies.Add objCookie.Name, objCookie
		End If
	End Sub

	Public Sub AddFromResponse( _
		ByVal objHttpResponse _
		)

		Me.AddFromResponseHeaders objHttpResponse.Headers
	End Sub

	Public Sub AddFromResponseHeaders( _
		ByVal varResponseHeaders _
		)

		If TypeName(varResponseHeaders) = "String" Then
			Me.AddFromString varResponseHeaders
		ElseIf TypeName(varResponseHeaders) = "base_HTTP_Headers" Then
			Me.AddFromString varResponseHeaders.ToString()
		End If
	End Sub

	Public Sub AddFromString( _
		ByVal strCookies _
		)

		Dim arrCookies, _
			strCookie, _
			intIndex
    
		arrCookies = Split(strCookies, vbCrLf)

		For intIndex = 0 To UBound(arrCookies)
			If arrCookies(intIndex) <> "" Then
				If Split(arrCookies(intIndex), ": ")(0) = "Set-Cookie" Then
					strCookie = Split(Split(arrCookies(intIndex), ": ")(1), "; ")(0)
					Me.Add strCookie
				End If
			End If
    		Next
	End Sub

	Public Sub Clear()
		p_objCookies.RemoveAll
	End Sub

	Public Sub ClearByDomain( _
		ByVal strDomain _
		)

		Dim varKey

		For Each varKey In p_objCookies.Keys()
			If p_objCookies(varKey).Domain = strDomain Then Me.Remove varKey
		Next
	End Sub

	Public Sub ClearByDomainPath( _
		ByVal strDomain, _
		ByVal strPath _
		)

		Dim varKey

		For Each varKey In p_objCookies.Keys()
			If p_objCookies(varKey).Domain = strDomain And p_objCookies(varKey).Path = strPath Then Me.Remove varKey
		Next
	End Sub

	Public Sub ClearSessionCookies()
		Dim varKey

		For Each varKey In p_objCookies.Keys()
			If p_objCookies(varKey).Expires = "" Then Me.Remove varKey
		Next
	End Sub

	Public Sub ClearTemporary()
		Dim varKey

		For Each varKey In p_objCookies.Keys()
			If p_objCookies(varKey).Expires = "" Or p_objCookies(varKey).Discard Then Me.Remove varKey
		Next
	End Sub

	Public Function Contains( _
		ByVal strName _
		)

		Contains = p_objCookies.Exists(strName)
	End Function

	Public Function FromResponse( _
		ByVal objResponse _
		)

		Me.FromResponseHeaders objResponse.Headers
	End Function

	Public Function FromResponseHeaders( _
    		ByVal varResponseHeaders _
		)

		If TypeName(varResponseHeaders) = "String" Then
			Me.FromString varResponseHeaders
		ElseIf TypeName(varResponseHeaders) = "base_HTTP_Headers" Then
			Me.FromString varResponseHeaders.ToString()
		End If
	End Function

	Public Function FromString( _
		ByVal strCookies _
		)

		With Me
			.Reset()
			.AddFromString strCookies
		End With
	End Function

	Public Function GetMatchingCookies( _
		ByVal strUrl _
		)

		Dim objArray
		Set objArray = New base_Data_Array

		If p_objCookies.Count > 0 Then
			Dim varKey

			For Each varKey In p_objCookies.Keys()
				If p_objCookies(varKey).Match(strUrl) Then objArray.Append p_objCookies(varKey)
			Next
		End If

		GetMatchingCookies = objArray.ToArray()
		Set objArray = Nothing
	End Function

	Public Function GetMatchingCookiesString( _
		ByVal strUrl _
		)

		Dim arrMatchingCookies, _
			strCookie

		arrMatchingCookies = Me.GetMatchingCookies(strUrl)

		strCookie = ""

		If UBound(arrMatchingCookies) >= 0 Then
			Dim intIndex

			strCookie = "Cookie: "

			For intIndex = 0 To UBound(arrMatchingCookies)
				strCookie = strCookie & arrMatchingCookies(intIndex).Name & "=" & arrMatchingCookies(intIndex).Value & "; "
			Next
        	
			strCookie = Left(strCookie, Len(strCookie) - 2)
		End If

		GetMatchingCookiesString = strCookie
	End Function

	Public Sub Load()
		With p_objTextStream
			.Open p_strFilename, FOR_READING
			Me.FromString .ReadAll()
		End With
	End Sub

	Public Sub Open( _
		ByVal strFilename _
		)

		p_strFilename = strFileName
		Me.Load 
	End Sub

	Public Sub Remove( _
    		ByVal strKey _
		)

		p_objCookies.Remove strKey
	End Sub

	Public Sub Reset()
		Class_Initialize()
	End Sub

	Public Sub Revert()
		Me.Load
	End Sub

	Public Sub Save()
		With p_objTextStream
			.Write ToCookieFileFormat()
			.Close
		End With
	End Sub

	Public Sub SaveAs( _
		ByVal strFilename _
		)

		p_objTextStream.Create strFileName
		Me.Save
		p_strFilename = strFileName
	End Sub

	Public Function ToDict()
		Set ToDict = p_objCookies
	End Function

	Public Function ToString()
		Dim strCookies

		strCookies = ""

		If p_objCookies.Count > 0 Then
			Dim varKey

			For Each varKey In p_objCookies.Keys()
				strCookies = strCookies & "Set-Cookie: " & p_objCookies(varKey).ToString() & vbCrLf
			Next

			strCookies = Left(strCookies, Len(strCookies) - 1)
		End If
    
		ToString = strCookies
	End Function

	
	' Helper Methods


	Private Function ToCookieFileFormat()
		Dim objCookie, _
			strCookieFile

		strCookieFile = ""

		For Each objCookie In p_objCookies.Items()
			With objCookie
				strCookieFile = strCookieFile & _
						.Domain & vbTab & _
			 			"TRUE" & vbTab & _
						.Path & vbTab & _
						UCase(.Secure) & vbTab & _
						.Expires & vbTab & _
						.Name & vbTab & _
						.Value & vbCrLf
			End With
		Next

		strCookieFile = Left(strCookieFile, Len(strCookieFile) - 1)

		ToCookieFileFormat = strCookieFile
	End Function

	Private Sub Class_Terminate()
		Set p_objCookies = Nothing
		Set p_objTextStream = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_HTTP_CookieJar.vbs" Then

End If
