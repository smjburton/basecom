Option Explicit

Function IIf(strExpression, varTrueResult, varFalseResult)
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
