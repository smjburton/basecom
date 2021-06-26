Option Explicit

Include "base_Data_Collection"
Include "base_HTTP_Header"

Class base_HTTP_Headers
	Private p_objHeaders

	Private Sub Class_Initialize()
		Set p_objHeaders = New base_Data_Collection
	End Sub


	' Properties


	Public Default Property Get Item( _
		ByVal intIndex _
		)
    
		Set Item = p_objHeaders.Item(intIndex)
	End Property

	Public Property Get Count()
		Count = p_objHeaders.Count
	End Property


	' Methods


	Public Sub AddHeaders( _
		ByVal objHeaders _
		)

		If TypeName(objHeaders) = "base_HTTP_Headers" Then
			Dim intIndex

			For intIndex = 0 To objHeaders.Count - 1
				Me.AddHeader objHeaders(intIndex)
			Next
		End If
	End Sub

	Public Sub AddHeader( _
		ByVal objHeader _
		)

		If TypeName(objHeader) = "base_HTTP_Header" Then p_objHeaders.Add objHeader
	End Sub

	Public Sub AddHeaderString( _
		ByVal strHeader _
		)

		If InStr(strHeader, ": ") > 0 Then
			Dim objHeader
			Set objHeader = New base_HTTP_Header

			objHeader.FromString strHeader
			Me.AddHeader objHeader
		End If
	End Sub

	Public Sub RemoveHeader( _
		ByVal strName _
		)

		Dim intHeaderIndex
		intHeaderIndex = IndexOf(strName)

		If IndexOf(strName) > 0 Then p_objHeaders.Remove intHeaderIndex
	End Sub

	Public Sub ClearHeaders()
		p_objHeaders.Clear
	End Sub

	Public Function IndexOf( _
		ByVal strName _
		)

		Dim intIndex

		For intIndex = 0 To p_objHeaders.Count - 1
			If p_objHeaders(intIndex).Name = strName Then
				IndexOf = intIndex
				Exit Function
			End If
		Next

		IndexOf = -1
	End Function

	Public Function GetHeader( _
		ByVal strName _
		)

		Dim intHeaderIndex
		intHeaderIndex = IndexOf(strName)

		If intHeaderIndex >= 0 Then
			Set GetHeader = p_objHeaders(intHeaderIndex)
		Else
			Set GetHeader = Nothing
		End If
	End Function

	Public Function HasHeader( _
		ByVal strName _
		)

		If IndexOf(strName) >= 0 Then
			HasHeader = True
		Else
			HasHeader = False
		End If
	End Function

	Public Sub FromString( _
		ByVal strHeaders _
		)
  
		p_objHeaders.Clear

		Dim arrHeaders, _
			intIndex
    
		arrHeaders = Split(strHeaders, vbCrLf)

		For intIndex = 0 To UBound(arrHeaders)
			Me.AddHeaderString arrHeaders(intIndex)
		Next
	End Sub

	Public Function ToString()
		Dim strHeaders, _
			intIndex

		strHeaders = ""

		For intIndex = 0 To p_objHeaders.Count - 1
			strHeaders = strHeaders & p_objHeaders(intIndex).ToString() & vbCrLf
		Next

		If Len(strHeaders) > 0 Then strHeaders = Left(strHeaders, Len(strHeaders) - 1)

		ToString = strHeaders
	End Function

	Public Function ToArray()
		ToArray = p_objHeaders.ToArray()
	End Function

	Public Function FromArray( _
		ByVal arrHeaders _
		)

		p_objHeaders.Clear

		Dim intIndex

		For intIndex = 0 To UBound(arrHeaders)
			Me.AddHeader arrHeaders(intIndex)
		Next
	End Function

	Private Sub Class_Terminate()
		Set p_objHeaders = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_HTTP_Headers.vbs" Then

End If