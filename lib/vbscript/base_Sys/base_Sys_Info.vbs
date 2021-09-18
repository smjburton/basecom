Option Explicit

Class base_Sys_Info
	Private p_objSysInfo

	Private Sub Class_Initialize()
		Set p_objSysInfo = CreateObject("ADSystemInfo")
	End Sub

	
	' Properties

	Public Property Get ComputerName()
		ComputerName = ""
	End Property
	
	Public Property Get DomainDnsName()
		DomainDnsName = ""
	End Property
		
	Public Property Get DomainShortName()
		DomainShortName = ""
	End Property

	Public Property Get ForestDnsName()
		ForestDnsName = ""
	End Property
		
	Public Property Get IsNativeMode()
		IsNativeMode = ""
	End Property
		
	Public Property Get PdcRoleOwner()
		PdcRoleOwner = ""
	End Property
	
	Public Property Get SchemaRoleOwner()
		SchemaRoleOwner = ""
	End Property
	
	Public Property Get SiteName()
		SiteName = ""
	End Property
	
	Public Property Get Username()
		Username = p_objSysInfo.UserName
	End Property


	' Methods


	Public Function GetAnyDcName()
		GetAnyDcName = ""
	End Function
	
	Public Function GetDcSiteName( _
		ByVal strServer _
		)

		GetDcSiteName = ""
	End Function
		
	Public Function GetTrees()
		GetTrees = ""
	End Function

	Public Sub RefreshSchemaCache()	

	End Sub


	Private Sub Class_Terminate()
		Set p_objSysInfo = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Sys_Info.vbs" Then

End If