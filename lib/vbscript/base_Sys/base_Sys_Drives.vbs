Option Explicit

Include "base_Sys_Drive"
Include "base_Data_Collection"

Class base_Sys_Drives
	Private p_Drives

	Private Sub Class_Initialize()
		Set p_Drives = New base_Data_Collection
	End Sub

	
	' Properties

	
	Public Property Get Count()
		Count = p_Drives.Count
	End Property

	Public Default Property Get Drive(intIndex)
		Set Drive = p_Drives(intIndex)
	End Property
	

	' Methods


	Public Sub FromDrives(objDrives)
		Dim objDrive, _
			objSysDrive

    		For Each objDrive in objDrives
			Set objSysDrive = New base_Sys_Drive
			objSysDrive.FromDrive objDrive
			p_Drives.Add objSysDrive 
    		Next		
	End Sub

	Private Sub Class_Terminate()
		Set p_Drives = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Sys_Drives.vbs" Then

End If