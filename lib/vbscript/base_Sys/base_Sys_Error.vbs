Option Explicit

Class base_Sys_Error
	Private p_intNumber, _
		p_strSource, _
		p_strDescription, _
		p_strHelpFile, _
		p_strHelpContext

	Private Sub Class_Initialize()
		' Err.Number
		' Err.Source
		' Err.Description
		' Err.HelpFile
		' Err.HelpContext
	End Sub

	Public Property Get Number()
		Number = p_intNumber
	End Property 

	Public Property Get Source()
		Source = p_strSource
	End Property 

	Public Property Get Description()
		Description = p_strDescription
	End Property 

	Public Property Get HelpFile()
		HelpFile = p_strHelpFile
	End Property 

	Public Property Get HelpContext()
		HelpContext = p_strHelpContext
	End Property 

	Public Default Function Init( _
		intErrorNumber, _
		strErrorSource, _
		strErrorDescription, _
		strErrorHelpFile, _
		strErrorHelpContext _
		)

		p_intNumber = intErrorNumber
		p_strSource = strErrorSource
		p_strDescription = strErrorDescription
		p_strHelpFile = strErrorHelpFile
		p_strHelpContext = strErrorHelpContext

		Set Init = Me
	End Function

	Private Sub Class_Terminate()

	End Sub
End Class

If WScript.ScriptName = "base_Sys_Error.vbs" Then

End If