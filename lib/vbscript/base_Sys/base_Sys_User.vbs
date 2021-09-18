Option Explicit

Include "base_Sys_Info"

Class base_Sys_User
	Private p_objSysInfo, _
		p_objUser

	Private Sub Class_Initialize()
		Set p_objSysInfo = New base_Sys_Info
		Set p_objSysInfo = GetObject("LDAP://" & objSysInfo.UserName)
 		' or specific user:
 		' Set objUser = GetObject("LDAP://CN=johndoe,OU=Users,DC=ss64,DC=com")
		' Default is to use the currently logged in user
	End Sub


	Public Property Get DistinguishedName
		DistinguishedName = objUser.distinguishedName
	End Property

	Public Property Get FirstName
FirstName = objUser.givenName
	End Property

	Public Property Get Initials
Initials = objUser.initials
	End Property

	Public Property Get LastName
LastName = objUser.sn
	End Property

	Public Property Get DisplayName
LastName = objUser.displayName
	End Property

	Public Property Get Description
Description = objUser.description
	End Property

	Public Property Get Office
 = objUser.physicalDeliveryOfficeName
	End Property

	Public Property Get TelephoneNumber
 = objUser.telephoneNumber
	End Property

	Public Property Get OtherTelephoneNumbers
 = objUser.otherTelephone
	End Property

	Public Property Get Email
 = objUser.mail
	End Property

	Public Property Get WebPage
 = objUser.wWWHomePage
	End Property

	Public Property Get OtherWebPages
 = objUser.url
	End Property

	Public Property Get StreetAddress
 = objUser.streetAddress
	End Property

	Public Property Get POBox
 = objUser.postOfficeBox
	End Property

	Public Property Get City
 = objUser.l
	End Property

	Public Property Get State
 = objUser.st
	End Property

	Public Property Get Province
 = objUser.st
	End Property

	Public Property Get ZipCode
 = objUser.postalCode
	End Property

	Public Property Get PostalCode
 = objUser.postalCode
	End Property

	Public Property Get Country
 = objUser.countryCode
	End Property

	Public Property Get UserLogonName
 = objUser.userPrincipalName
	End Property

	Public Property Get PreWindows2000LogonName
 = objUser.sAMAccountName
	End Property

	Public Property Get AccountDisabled
 = objUser.AccountDisabled
	End Property

	Public Property Get LogonHours
 = objUser.logonHours
	End Property

	Public Property Get LogonWorkstations
 = objUser.userWorkstations
	End Property

	Public Property Get UserAccountControl
 = objUser.userAccountControl
	End Property

	Public Property Get ProfilePath
 = objUser.profilePath
	End Property

	Public Property Get LogonScriptPath
 = objUser.scriptPath
	End Property

	Public Property Get HomeDirectory
 = objUser.homeDirectory
	End Property

	Public Property Get HomeDrive
 = objUser.homeDrive
	End Property

	Public Property Get HomePhone
 = objUser.homePhone
	End Property

	Public Property Get OtherHomePhone
 = objUser.otherHomePhone
	End Property

	Public Property Get Pager
 = objUser.pager
	End Property

	Public Property Get OtherPager
 = objUser.otherPager
	End Property

	Public Property Get Mobile
 = objUser.mobile
	End Property

	Public Property Get OtherMobile
 = objUser.otherMobile
	End Property

	Public Property Get Fax
 = objUser.facsimileTelephoneNumber
	End Property

	Public Property Get OtherFax
 = objUser.otherFacsimileTelephoneNumber
	End Property

	Public Property Get IpPhone
 = objUser.ipPhone
	End Property

	Public Property Get OtherIpPhone
 = objUser.otherIpPhone
	End Property

	Public Property Get Notes
 = objUser.info
	End Property

	Public Property Get DepartmentTitle
 = objUser.title
	End Property

	Public Property Get Department
 = objUser.department
	End Property

	Public Property Get Company
 = objUser.company
	End Property

	Public Property Get Manager
 = objUser.manager
	End Property



' WScript.Network.1

' ComputerName	Property ComputerName As String
'     read-only
' UserDomain	Property UserDomain As String
'     read-only
' UserName	Property UserName As String
'     read-only

' ADSystemInfo

' ExpandEnvironmentStrings	Function ExpandEnvironmentStrings(Src As String) As String

' WScript.Shell.Environment

' Environment
' Valid parameters for Environment are PROCESS, SYSTEM, USER and VOLATILE.

	Private Sub Class_Terminate()
		Set p_objSysInfo = Nothing
		Set p_objSysInfo = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Sys_User.vbs" Then

End If
