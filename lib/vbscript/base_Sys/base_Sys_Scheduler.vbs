Option Explicit

Class base_Sys_Scheduler
	Private p_Scheduler

	Private Sub Class_Initialize()
		Set p_Scheduler = CreateObject("Schedule.Service")
	End Sub


	' Properties


	Public Property Get Connected()
		Connected = p_Scheduler.Connected 
	End Property

	Public Property Get ConnectedDomain()
		ConnectedDomain = p_Scheduler.ConnectedDomain 
	End Property

	Public Property Get ConnectedUser()
		ConnectedDomain = p_Scheduler.ConnectedUser 
	End Property
  
	Public Property Get HighestVersion()
		If IsObject(p_Scheduler.HighestVersion) Then
			Set HighestVersion = p_Scheduler.HighestVersion
		Else
			HighestVersion = p_Scheduler.HighestVersion
		End If
	End Property

	Public Default Property Get TargetServer()
		TargetServer = p_Scheduler.TargetServer
	End Property


	' Methods


	Public Sub Connect() ' Optional params: [serverName], [user], [domain], [password])
		p_Scheduler.Connect
	End Sub

	Public Function GetFolder(strPath)
		Set GetFolder = p_Scheduler.GetFolder(strPath)
	End Function

	Public Function GetRunningTasks(lngFlags)
		Set GetRunningTasks = p_Scheduler.GetRunningTasks(lngFlags)
	End Function

	Public Function NewTask(varFlags)
		Set NewTask = p_Scheduler.NewTask(varFlags)
	End Function

	Private Sub Class_Terminate()
		Set p_Scheduler = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Scheduler.vbs" Then

End If
