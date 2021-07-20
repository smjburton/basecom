Option Explicit

Class base_Peripheral_Fax_TIFF
	Private p_FaxTiff

	Private Sub Class_Initialize()
		Set p_FaxTiff = CreateObject("FaxTiff.FaxTiff.1")
	End Sub


	' Properties




	' Methods



	Private Sub Class_Terminate()
		Set p_FaxTiff = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Peripheral_Fax_TIFF.vbs" Then

End If