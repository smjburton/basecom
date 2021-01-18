Option Explicit

Class base_Fax_Controller
	Private p_FaxControl

	Private Sub Class_Initialize()
		Set p_FaxControl = CreateObject("FaxControl.FaxControl.1")
	End Sub


	' Properties


	Public Property Get IsFaxServiceInstalled()
		IsFaxServiceInstalled = p_FaxControl.IsFaxServiceInstalled
	End Property

	Public Property Get IsLocalFaxPrinterInstalled()
		IsLocalFaxPrinterInstalled = p_FaxControl.IsLocalFaxPrinterInstalled
	End Property


	' Methods


	Public Sub InstallFaxService()
		p_FaxControl.InstallFaxService
	End Sub

	Public Sub InstallLocalFaxPrinter()
		p_FaxControl.InstallLocalFaxPrinter
	End Sub

	Private Sub Class_Terminate()
		Set p_FaxControl = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Fax_Controller.vbs" Then

End If