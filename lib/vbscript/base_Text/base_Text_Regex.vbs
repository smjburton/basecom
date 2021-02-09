Option Explicit

Class base_Text_Regex
	Private p_Regex

	Private Sub Class_Initialize()
		Set p_Regex = CreateObject("VBScript.RegExp")
	End Sub


	' Properties


	Public Property Get Global()
		Global = p_Regex.Global
	End Property

	Public Property Let Global(blnGlobal)
		p_Regex.Global = blnGlobal
	End Property

	Public Property Get IgnoreCase()
		IgnoreCase = p_Regex.IgnoreCase
	End Property

	Public Property Let IgnoreCase(blnIgnoreCase)
		p_Regex.IgnoreCase = blnIgnoreCase
	End Property

	Public Property Get Multiline()
		Multiline = p_Regex.Multiline
	End Property

	Public Property Let Multiline(blnMultiline)
		p_Regex.Multiline = blnMultiline
	End Property

	Public Property Get Pattern()
		Pattern = p_Regex.Pattern
	End Property

	Public Property Let Pattern(strPattern)
		p_Regex.Pattern = strPattern
	End Property


	' Methods


	Public Function Execute(strSourceString)
		Set Execute = p_Regex.Execute(strSourceString)
	End Function

	Public Function Replace(strSourceString, varReplace)
		Replace = p_Regex.Replace(strSourceString, varReplace)
	End Function

	Public Function Test(strSourceString)
		Test = p_Regex.Test(strSourceString)
	End Function

	Private Sub Class_Terminate()
		Set p_Regex = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Text_Regex.vbs" Then

End If