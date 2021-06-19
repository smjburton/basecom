Option Explicit

Class base_HTTP_Header
	Private p_strName, _
        	p_strValue

	Private Sub Class_Initialize()

	End Sub


	' Properties


	Public Property Get Name()
		Name = p_strName
	End Property

	Public Property Get Value()
		Value = p_strValue
	End Property


	' Methods	


	Public Sub FromString( _
		ByVal strHeader _
		)

		p_strName = ExtractHeaderName(strHeader)
		p_strValue = ExtractHeaderValue(strHeader)
	End Sub

	Public Function ToString()
		ToString = p_strName & ": " & p_strValue
	End Function


	' Helper Functions


	Private Function ExtractHeaderName( _
		ByVal strHeader _
		)

		ExtractHeaderName = Split(strHeader, ": ")(0)
	End Function

	Private Function ExtractHeaderValue( _
		ByVal strHeader _
		)

		ExtractHeaderValue = Split(Split(strHeader, ": ")(1), "; ")(0)
	End Function
End Class

If WScript.ScriptName = "base_HTTP_Header.vbs" Then

End If
