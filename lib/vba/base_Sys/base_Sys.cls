Option Explicit

Sub Print(ByVal strText)
    WScript.Echo strText
End Sub

Sub Sleep(ByVal intTimeSeconds)
    WScript.Sleep Int(intTimeSeconds * 1000)
End Sub

