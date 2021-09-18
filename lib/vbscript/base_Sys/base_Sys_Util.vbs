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

' Function FunctionExists( func_name )'
'     FunctionExists = False 
' 
'     On Error Resume Next
' 
'     Dim f : Set f = GetRef(func_name)
' 
'     If Err.number = 0 Then
'         FunctionExists = True
'     End If  
'     On Error GoTo 0
' 
' End Function

If WScript.ScriptName = "base_Sys_Util.vbs" Then

End If
