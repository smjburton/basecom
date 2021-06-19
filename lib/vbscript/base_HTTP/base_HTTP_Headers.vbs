Option Explicit

Include "base_Data_Collection"
Include "base_HTTP_Header"

Class base_HTTP_Headers
	Private p_objHeaders

	Private Sub Class_Initialize()
		Set p_objHeaders = New base_Data_Collection
	End Sub


	' Properties


	Public Property Get Item( _
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

		Dim intIndex
    
		If TypeName(objHeader) = "base_HTTP_Headers" Then
			For intIndex = 1 To objHeaders.Count
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
    
		Dim objHeader
		Set objHeader = New base_HTTP_Header

		objHeader.FromString strHeader
		Me.AddHeader objHeader
	End Sub

	Public Sub RemoveHeader( _
		ByVal strName _
		)

	End Sub

	Public Sub ClearHeaders()
		Dim intIndex

		For intIndex = 1 To p_objHeaders.Count
			p_objHeaders.Remove 1
		Next
	End Sub

	Public Sub GetHeader( _
		ByVal strName _
		)

	End Sub

	Public Function HasHeader( _
		ByVal strName _
		)

	End Function

	Public Sub FromString( _
		ByVal strHeaders _
		)
  
		Dim arrHeaders As Variant, _
			intIndex As Integer
    
		arrHeaders = Split(strHeaders, vbCrLf)
    
		For intIndex = 0 To UBound(arrHeaders) - 2
			Me.AddHeaderString arrHeaders(intIndex)
		Next
	End Sub

	Public Function ToString()
		Dim strHeaders, _
			intIndex As Integer

		strHeaders = ""

		For intIndex = 1 To p_objHeaders.Count
			strHeaders = strHeaders & p_objHeaders(intIndex).ToString() & vbCrLf
		Next
    
		ToString = strHeaders
	End Function

	Public Function ToArray()

	End Function

	Public Function FromArray()

	End Function

	Private Sub Class_Terminate()
		Set p_objHeaders = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_HTTP_Headers.vbs" Then

End If
