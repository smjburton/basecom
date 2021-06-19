Option Explicit

Class base_HTTP_Session
	Private Sub Class_Initialize()

	End Sub


	' Properties


	Public Property Get DefaultHeader()

	End Property

	Public Sub DefaultHeaders(arrHeaders)

	End Sub

	Public Property Get KeepAlive()
		KeepAlive = pKeepAlive
	End Property

	Public Property Let KeepAlive(blnKeepAlive)
		pKeepAlive = blnKeepAlive
	End Property

	Public Property Get StoreCookies()
		StoreCookies = pStoreCookies
	End Property

	Public Property Let StoreCookies(blnStoreCookies)
		pStoreCookies = blnStoreCookies
	End Property


	' Methods


	Public Function Request()

	End Function

	Public Sub Send()

	End Sub

	Public Sub ClearDefaultHeaders()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

If WScript.ScriptName = "base_HTTP_Session.vbs" Then

End If
