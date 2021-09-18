Option Explicit

Class base_Sys_Admin
	Private p_objSysAdmin

	Private Sub Class_Initialize()

	End Sub


	' Methods


	Public Sub CreateUser()
Dim strUserName
Dim objRootLDAP
Dim objContainer
Dim objNewUser
strUserName = "MorganTestUser" 
 
Set objRootLDAP = GetObject("LDAP://rootDSE")
 
' You can give your own OU like <i>LDAP://OU=TestOU</i> instead of <i>LDAP://CN=Users</i>
Set objContainer = GetObject("LDAP://CN=Users," & _
objRootLDAP.Get("defaultNamingContext")) 
 
Set objNewUser = objContainer.Create("User", "cn=" & strUserName)
objNewUser.Put "sAMAccountName", strUserName
objNewUser.Put "givenName", "Morgan"
objNewUser.Put "sn", "TestUser"
objNewUser.Put "displayName", "Morgan TestUser"
objNewUser.Put "Description", "AD User created by VB Script"
objNewUser.SetInfo
 
objNewUser.SetPassword "MyPassword123"
objNewUser.Put "PasswordExpired", CLng(1)
objNewUser.AccountDisabled = FALSE
 
MsgBox ("New Active Directory User created successfully by using VB Script...")
 
WScript.Quit

' With user input:
Dim strUserName
Dim objRootLDAP
Dim objContainer
Dim objNewUser
 
Do
   strUserName = InputBox ("Please enter user name")
   If strUserName = "" then
      Msgbox "No user name entered"
   end if
Loop Until strUserName <> ""
 
MsgBox "Please click OK to continue..."
 
Set objRootLDAP = GetObject("LDAP://rootDSE")
 
' You can give your own OU like <i>LDAP://OU=TestOU</i> instead of <i>LDAP://CN=Users</i>
Set objContainer = GetObject("LDAP://CN=Users," & _
objRootLDAP.Get("defaultNamingContext")) 
 
Set objNewUser = objContainer.Create("User", "cn=" & strUserName)
objNewUser.Put "sAMAccountName", strUserName
objNewUser.Put "Description", "AD User created by VB Script"
objNewUser.SetInfo
 
objNewUser.SetPassword "MyPassword123"
objNewUser.Put "PasswordExpired", CLng(1)
objNewUser.AccountDisabled = FALSE
 
MsgBox ("New Active Directory User created successfully by using VB Script...")
 
WScript.Quit
	End Sub

	Public Sub CreateUsers()
Dim strUserName
Dim objRootLDAP
Dim objContainer
Dim objNewUser
Dim i
 
For i = 0 To 5
 
strUserName = "MorganTestUser"& i 
 
Set objRootLDAP = GetObject("LDAP://rootDSE")
 
' You can give your own OU like <i>LDAP://OU=TestOU</i> instead of <i>LDAP://CN=Users</i>
Set objContainer = GetObject("LDAP://CN=Users," & _
objRootLDAP.Get("defaultNamingContext")) 
 
 
Set objNewUser = objContainer.Create("User", "cn=" & strUserName)
objNewUser.Put "sAMAccountName", strUserName
objNewUser.Put "Description", "Bulk AD User created by VB Script"
objNewUser.SetInfo
 
objNewUser.SetPassword "MyPassword123"
objNewUser.Put "PasswordExpired", CLng(1)
objNewUser.AccountDisabled = FALSE
 
Next
 
MsgBox (i &" AD Users created successfully by using VB Script...")
	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

If WScript.ScriptName = "base_Sys_Admin.vbs" Then

End If