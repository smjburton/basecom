Option Explicit

Function IIf( _
	ByVal strExpression, _
	ByVal varTrueResult, _
	ByVal varFalseResult _
	)

	If strExpression Then
		If IsObject(varTrueResult) Then
			Set IIf = varTrueResult
		Else
			IIf = varTrueResult
		End If
	Else
		If IsObject(varFalseResult) Then
			Set IIf = varFalseResult
		Else
			IIf = varFalseResult
		End If
	End If
End Function

If WScript.ScriptName = "base_Sys_Util.vbs" Then

End If
