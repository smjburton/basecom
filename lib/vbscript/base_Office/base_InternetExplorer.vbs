Option Explicit

Class base_InternetExplorer
	Private p_InternetExplorer

	Private Sub Class_Initialize()
		Set p_InternetExplorer = CreateObject("InternetExplorer.Application")
	End Sub


	' Properties


	Public Property Get AddressBar()
		AddressBar = p_InternetExplorer.AddressBar 
	End Property

	Public Property Let AddressBar(blnAddressBar)
		p_InternetExplorer.AddressBar = blnAddressBar
	End Property

	Public Property Get Application()
		Set Application = p_InternetExplorer.Application		
	End Property

	Public Property Get Busy()
		Busy = p_InternetExplorer.Busy 		
	End Property

	Public Property Get Container()
		Set Container = p_InternetExplorer.Container
	End Property

	Public Property Get Document()
		Set Document = p_InternetExplorer.Document
	End Property

	Public Property Get FullName()
		FullName = p_InternetExplorer.FullName
	End Property

	Public Property Get FullScreen()
		FullScreen = p_InternetExplorer.FullScreen
	End Property

	Public Property Let FullScreen(blnFullScreen) 
		p_InternetExplorer.FullScreen = blnFullScreen
	End Property

	Public Property Get Height() 
		Height = p_InternetExplorer.Height 
	End Property

	Public Property Let Height(lngHeight) 
		p_InternetExplorer.Height = lngHeight
	End Property

	Public Property Get HWND()
		HWND = p_InternetExplorer.HWND
	End Property

	Public Property Get Left()
		Left = p_InternetExplorer.Left
	End Property

	Public Property Let Left(lngLeft) 
		p_InternetExplorer.Left = lngLeft
	End Property

	Public Property Get LocationName()
		LocationName = p_InternetExplorer.LocationName 
	End Property

	Public Property Get LocationURL()
		LocationURL = p_InternetExplorer.LocationURL 
	End Property

	Public Property Get MenuBar()
		MenuBar = p_InternetExplorer.MenuBar
	End Property

	Public Property Let MenuBar(blnMenuBar) 
		p_InternetExplorer.MenuBar = blnMenuBar
	End Property

	Public Property Get Name()
		Name = p_InternetExplorer.Name 
	End Property

	Public Property Get Offline() 
		Offline = p_InternetExplorer.Offline 
	End Property

	Public Property Let Offline(blnOffline) 
		p_InternetExplorer.Offline = blnOffline
	End Property

	Public Property Get Parent()
		Set Parent = p_InternetExplorer.Parent
	End Property

	Public Property Get Path()
		Path = p_InternetExplorer.Path
	End Property

	Public Property Get ReadyState()
		Set ReadyState = p_InternetExplorer.ReadyState 
	End Property

	Public Property Get RegisterAsBrowser()
		RegisterAsBrowser = p_InternetExplorer.RegisterAsBrowser
	End Property

	Public Property Let RegisterAsBrowser(blnRegisterAsBrowser)
		p_InternetExplorer.RegisterAsBrowser = blnRegisterAsBrowser
	End Property

	Public Property Get RegisterAsDropTarget()
		RegisterAsDropTarget = p_InternetExplorer.RegisterAsDropTarget
	End Property

	Public Property Let RegisterAsDropTarget(blnRegisterAsDropTarget)
		p_InternetExplorer.RegisterAsDropTarget = blnRegisterAsDropTarget
	End Property

	Public Property Get Resizable()
		Resizable = p_InternetExplorer.Resizable
	End Property
 
	Public Property Let Resizable(blnResizable)
		p_InternetExplorer.Resizable = blnResizable
	End Property

	Public Property Get Silent() 
		Silent = p_InternetExplorer.Silent
	End Property

	Public Property Let Silent(blnSilent) 
		p_InternetExplorer.Silent = blnSilent
	End Property

	Public Property Get StatusBar()
		StatusBar = p_InternetExplorer.StatusBar
	End Property

	Public Property Let StatusBar(blnStatusBar) 
		p_InternetExplorer.StatusBar = blnStatusBar
	End Property

	Public Property Get StatusText()
		StatusText = p_InternetExplorer.StatusText
	End Property

	Public Property Let StatusText(strStatusText)
		p_InternetExplorer.StatusText = strStatusText
	End Property

	Public Property Get TheaterMode()
		TheaterMode = p_InternetExplorer.TheaterMode
	End Property

	Public Property Let TheaterMode(blnTheaterMode)
		p_InternetExplorer.TheaterMode = blnTheaterMode
	End Property

	Public Property Get ToolBar()
		ToolBar = p_InternetExplorer.ToolBar
	End Property

	Public Property Let ToolBar(lngToolBar)
		p_InternetExplorer.ToolBar = lngToolBar
	End Property

	Public Property Get Top()
		Top = p_InternetExplorer.Top
	End Property

	Public Property Let Top(lngTop)
		p_InternetExplorer.Top = lngTop
	End Property

	Public Property Get TopLevelContainer()
		TopLevelContainer = p_InternetExplorer.TopLevelContainer
	End Property

	Public Property Get DocType()
		DocType = p_InternetExplorer.Type
	End Property

	Public Property Get Visible()
		Visible = p_InternetExplorer.Visible
	End Property

	Public Property Let Visible(blnVisible) 
		p_InternetExplorer.Visible = blnVisible
	End Property

	Public Property Get Width()
		Width = p_InternetExplorer.Width
	End Property

	Public Property Let Width(lngWidth) 
		p_InternetExplorer.Width = lngWidth
	End Property

	
	' Methods


	Public Sub ClientToWindow(lngPcx, lngPcy)
		p_InternetExplorer.ClientToWindow lngPcx, lngPcy
	End Sub

	Public Sub ExecWB(objCmdID, objCmdExecOpt) ' Optional Params: [pvaIn], [pvaOut]
		p_InternetExplorer.ExecWB objCmdID, objCmdExecOpt
	End Sub

	Public Function GetProperty(strProperty)
		p_InternetExplorer.GetProperty strProperty
	End Function
 
	Public Sub GoBack() 
		p_InternetExplorer.GoBack
	End Sub

	Public Sub GoForward() 
		p_InternetExplorer.GoForward
	End Sub

	Public Sub GoHome()
		p_InternetExplorer.GoHome
	End Sub
 
	Public Sub GoSearch() 
		p_InternetExplorer.GoSearch
	End Sub

	Public Sub Navigate(strUrl)  ' Optional params: [Flags], [TargetFrameName], [PostData], [Headers]
		p_InternetExplorer.Navigate strUrl
	End Sub

	Public Sub Navigate2(objUrl) ' Optional params: [Flags], [TargetFrameName], [PostData], [Headers]) 
		p_InternetExplorer.Navigate2 objUrl
	End Sub

	Public Sub PutProperty(strProperty, objVtValue) 
		p_InternetExplorer.PutProperty strProperty, objVtValue
	End Sub

	Public Function QueryStatusWB(objCmdId)
		Set QueryStatusWB = p_InternetExplorer.QueryStatusWB(objCmdId)
	End Function
 
	Public Sub Quit() 
		p_InternetExplorer.Quit
	End Sub

	Public Sub Refresh() 
		p_InternetExplorer.Refresh
	End Sub

	Public Sub Refresh2() ' Optional params: [Level]
		p_InternetExplorer.Refresh2
	End Sub

	Public Sub ShowBrowserBar(objPvaClsId) ' Optional params: [pvarShow], [pvarSize]
		p_InternetExplorer.ShowBrowserBar objPvaClsId
	End Sub

	Public Sub Terminate() 
		p_InternetExplorer.Stop
	End Sub

	Private Sub Class_Terminate()
		Set p_InternetExplorer = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_InternetExplorer.vbs" Then
	Dim explorer
	Set explorer = New base_InternetExplorer

	With explorer
		.Visible = True
		.Navigate "http://www.google.com"

		Do While .Busy
			WScript.Sleep 200
		Loop

		WScript.Echo CStr(.Document.DocumentElement.innerText)

		.Quit
	End With

	Set explorer = Nothing
End If
