Option Explicit

Class base_Peripheral_Fax_Document
	Private p_FaxDocument

	Private Sub Class_Initialize()
		Set p_FaxDocument = CreateObject("FaxComEx.FaxDocument.1")
	End Sub


	' Properties




	' Methods



	Private Sub Class_Terminate()
		Set p_FaxDocument = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Peripheral_Fax_Document.vbs" Then

End If