Option Explicit

Class vbs_Script
	Private pScriptHost, _
		pScriptEngine, _
		pScriptLanguage

	Private Sub Class_Initialize()
		Set pScriptHost = CreateObject("HTMLFile")
		Set pScriptEngine = pScriptHost.parentWindow
		pScriptLanguage = ""
	End Sub


	' Properties:


	Public Property Get Error()

	End Property

	Public Property Get Language()
		Language = pScriptLanguage
	End Property

	Public Property Let Language(strLang)
		If LCase(strLang) = "vbscript" Or LCase(strLang) = "jscript" Then pScriptLanguage = strLang
	End Property

	Public Property Get Variable(strVar)
		If Exists(strVar) Then
			If TypeName(Eval("pScriptEngine." & strVar)) = "JScriptTypeInfo" Then
				Set Variable = Eval("pScriptEngine." & strVar)
			Else
				Variable = Eval("pScriptEngine." & strVar)
			End If
		End If
	End Property

	Public Property Let Variable(strVar, strNewVal)
		If Not Exists(strVar) Then	
			If LCase(pScriptLanguage) = "jscript" Then
				pScriptEngine.execScript "var " & strVar & ";", pScriptLanguage
			ElseIf LCase(pScriptLanguage) = "vbscript" Then
				pScriptEngine.execScript "Dim " & strVar, pScriptLanguage
			End If
		End If

		If TypeName(strNewVal) = "String" And Left(strNewVal, 1) = "{" And Right(strNewVal, 1) = "}" Then
			pScriptEngine.execScript strVar & " = " & strNewVal, pScriptLanguage
		Else
			Execute("pScriptEngine." & strVar & " = strNewVal")
		End If
	End Property

	Public Property Set Variable(strVar, objNewObj)
		If Not Exists(strVar) Then
			If LCase(pScriptLanguage) = "jscript" Then
				pScriptEngine.execScript "var " & strVar & ";", pScriptLanguage
			ElseIf LCase(pScriptLanguage) = "vbscript" Then
				pScriptEngine.execScript "Dim " & strVar, pScriptLanguage
			End If
		End If

		Execute("Set pScriptEngine." & strVar & " = objNewObj")
	End Property

	Public Property Get VarType(strVar)

	End Property


	' Methods:


	Public Sub AddCode(strCode)
		On Error Resume Next

		If pScriptLanguage <> "" Then pScriptEngine.execScript strCode, pScriptLanguage

		If Err.Number <> 0 Then
			WScript.Echo "Error occured in 'AddCode()'."
		End If
	End Sub

	Public Function Exists(strVar)
		On Error Resume Next

		Eval("pScriptEngine." & strVar)

		If Err.Number <> 0 Then
			Err.Clear
			Exists = False
		Else
			Exists = True
		End If
	End Function

	Public Function Run(strProcedure, arrArgs)
		Dim strProc, _
			strArgs, _
			i

		If Right(strProcedure, 2) = "()" Then strProcedure = Left(strProcedure, Len(strProcedure) - 2)

		strArgs = "("

		If IsArray(arrArgs) Then
			If UBound(arrArgs) >= 0 Then
				For i = 0 to UBound(arrArgs)
					strArgs = strArgs & "arrArgs(" & i & "), "
				Next

				strArgs = Left(strArgs, Len(strArgs) - 2)
			End If
		End If

		strArgs = strArgs & ")"

		If TypeName(Eval("pScriptEngine." & strProcedure & strArgs)) = "JScriptTypeInfo" Then
			Set Run = Eval("pScriptEngine." & strProcedure & strArgs)
		Else
			Run = Eval("pScriptEngine." & strProcedure & strArgs)
		End If
	End Function

	Public Sub Reset()
		Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
		Set pScriptHost = Nothing
		Set pScriptEngine = Nothing
	End Sub
End Class

If WScript.ScriptName = "vbs_Script.vbs" Then
	Dim script
	Set script = New vbs_Script

	With script
		.Language = "JScript"
		.AddCode("function addNumbers(i, j) { return i + j; }")
		.Variable("test") = "Hello, world!"
	End With

	WScript.Echo script.Run("addNumbers", Array(1, 3))
	WScript.Echo script.Variable("test")
End If