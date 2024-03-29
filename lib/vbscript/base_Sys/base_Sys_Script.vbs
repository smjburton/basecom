Option Explicit

Class base_Sys_Script
	Private p_objScriptHost, _
		p_objScriptEngine, _
		p_strScriptLanguage

	Private Sub Class_Initialize()
		Set p_objScriptHost = CreateObject("HTMLFile")
		Set p_objScriptEngine = p_objScriptHost.parentWindow
		p_strScriptLanguage = ""
	End Sub


	' Properties:


	Public Property Get Error()

	End Property

	Public Property Get Language()
		Language = p_strScriptLanguage
	End Property

	Public Property Let Language(strLang)
		If LCase(strLang) = "vbscript" Or LCase(strLang) = "jscript" Then p_strScriptLanguage = strLang
	End Property

	Public Property Get Variable(strVar)
		If Exists(strVar) Then
			If TypeName(Eval("p_objScriptEngine." & strVar)) = "JScriptTypeInfo" Then
				Set Variable = Eval("p_objScriptEngine." & strVar)
			Else
				Variable = Eval("p_objScriptEngine." & strVar)
			End If
		End If
	End Property

	Public Property Let Variable(strVar, strNewVal)
		If Not Exists(strVar) Then	
			If LCase(p_strScriptLanguage) = "jscript" Then
				p_objScriptEngine.execScript "var " & strVar & ";", p_strScriptLanguage
			ElseIf LCase(p_strScriptLanguage) = "vbscript" Then
				p_objScriptEngine.execScript "Dim " & strVar, p_strScriptLanguage
			End If
		End If

		If TypeName(strNewVal) = "String" And Left(strNewVal, 1) = "{" And Right(strNewVal, 1) = "}" Then
			p_objScriptEngine.execScript strVar & " = " & strNewVal, p_strScriptLanguage
		Else
			Execute("p_objScriptEngine." & strVar & " = strNewVal")
		End If
	End Property

	Public Property Set Variable(strVar, objNewObj)
		If Not Exists(strVar) Then
			If LCase(p_strScriptLanguage) = "jscript" Then
				p_objScriptEngine.execScript "var " & strVar & ";", p_strScriptLanguage
			ElseIf LCase(p_strScriptLanguage) = "vbscript" Then
				p_objScriptEngine.execScript "Dim " & strVar, p_strScriptLanguage
			End If
		End If

		Execute("Set p_objScriptEngine." & strVar & " = objNewObj")
	End Property

	Public Property Get VarType(strVar)

	End Property


	' Methods:


	Public Sub AddCode(strCode)
		On Error Resume Next

		If p_strScriptLanguage <> "" Then p_objScriptEngine.execScript strCode, p_strScriptLanguage

		If Err.Number <> 0 Then
			WScript.Echo "Error occured in 'AddCode()'."
		End If
	End Sub

	Public Function Exists(strVar)
		On Error Resume Next

		Eval("p_objScriptEngine." & strVar)

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

		If TypeName(Eval("p_objScriptEngine." & strProcedure & strArgs)) = "JScriptTypeInfo" Then
			Set Run = Eval("p_objScriptEngine." & strProcedure & strArgs)
		Else
			Run = Eval("p_objScriptEngine." & strProcedure & strArgs)
		End If
	End Function

	Public Sub Reset()
		Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
		Set p_objScriptHost = Nothing
		Set p_objScriptEngine = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Sys_Script.vbs" Then
	Dim objScript
	Set objScript = New base_Sys_Script

	With objScript
		.Language = "JScript"
		.AddCode("function addNumbers(i, j) { return i + j; }")
		.Variable("test") = "Hello, world!"
	End With

	WScript.Echo objScript.Run("addNumbers", Array(1, 3))
	WScript.Echo objScript.Variable("test")
End If
