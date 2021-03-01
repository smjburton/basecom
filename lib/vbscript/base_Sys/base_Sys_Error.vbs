Option Explicit

Class base_Sys_Error
	Private p_Number, _
		p_Source, _
		p_Description, _
		p_HelpFile, _
		p_HelpContext

	Private Sub Class_Initialize()
		' Err.Number
		' Err.Source
		' Err.Description
		' Err.HelpFile
		' Err.HelpContext
	End Sub

	Public Property Get Number()
		Number = p_Number
	End Property 

	Public Property Get Source()
		Source = p_Source
	End Property 

	Public Property Get Description()
		Description = p_Description
	End Property 

	Public Property Get HelpFile()
		HelpFile = p_HelpFile
	End Property 

	Public Property Get HelpContext()
		HelpContext = p_HelpContext
	End Property 

	Public Default Function Init( _
		intErrorNumber, _
		strErrorSource, _
		strErrorDescription, _
		strErrorHelpFile, _
		strErrorHelpContext _
		)

		p_Number = intErrorNumber
		p_Source = strErrorSource
		p_Description = strErrorDescription
		p_HelpFile = strErrorHelpFile
		p_HelpContext = strErrorHelpContext

		Set Init = Me
	End Function

	Private Sub Class_Terminate()

	End Sub
End Class

If WScript.ScriptName = "base_Sys_Error.vbs" Then

End If