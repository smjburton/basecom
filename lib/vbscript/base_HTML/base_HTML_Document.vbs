Option Explicit

Sub Include(file)
	Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")
	ExecuteGlobal FSO.OpenTextFile(file & ".vbs", 1).ReadAll()
	Set FSO = Nothing
End Sub

Include "v_Data_Array"

Class base_HTML_Document
	Private pHtmlDoc


	' Constructor


	Private Sub Class_Initialize()
		Set pHtmlDoc = CreateObject("HTMLFile")
	End Sub


	' Properties


	Public Property Get ActiveElement()
		Set ActiveElement = pHtmlDoc.activeElement
	End Property

	Public Property Get ALinkColor()
		Set ALinkColor = pHtmlDoc.alinkColor
	End Property

	Public Property Get All()
		Set All = pHtmlDoc.all
	End Property

	Public Property Get Anchors()
		Set Anchors = pHtmlDoc.anchors
	End Property

	Public Property Get Applets()
		Set Applets = pHtmlDoc.applets
	End Property

	Public Property Get Attributes()
		Set Attributes = pHtmlDoc.attributes
	End Property

	Public Property Get BgColor()
		BgColor = pHtmlDoc.bgColor
	End Property

	Public Property Get Body()
		Set Body = pHtmlDoc.body
	End Property

	Public Property Get CharacterSet()
		CharacterSet = pHtmlDoc.charset
	End Property

	Public Property Get ChildNodes()
		Set ChildNodes = pHtmlDoc.childNodes
	End Property

	Public Property Get Compatible()
		Set Compatible = pHtmlDoc.compatible
	End Property

	Public Property Get CompatMode()
		CompatMode = pHtmlDoc.compatMode
	End Property

	Public Property Get ContentType()
		On Error Resume Next
		Dim strContentType
		strContentType = pHtmlDoc.mimeType
		If Err.Number = 0 Then ContentType = strContentType
	End Property

	Public Property Get Cookie()
		Cookie = pHtmlDoc.cookie
	End Property

	Public Property Get DefaultCharset()
		DefaultCharset = pHtmlDoc.defaultCharset
	End Property

	Public Property Get DefaultView()
		Set DefaultView = pHtmlDoc.parentWindow
	End Property

	Public Property Get DesignMode()
		DesignMode = pHtmlDoc.designMode
	End Property

	Public Property Get Dir()
		Dir = pHtmlDoc.dir
	End Property

	Public Property Get DocType()
		Set DocType = pHtmlDoc.doctype
	End Property

	Public Property Get DocumentElement()
		Set DocumentElement = pHtmlDoc.documentElement
	End Property

	Public Property Get DocumentMode()
		Set DocumentMode = pHtmlDoc.documentMode
	End Property

	Public Property Get DocumentURI()
		DocumentURI = pHtmlDoc.url
	End Property

	Public Property Get Domain()
		Domain = pHtmlDoc.domain
	End Property

	Public Property Get Embeds()
		Embeds = pHtmlDoc.embeds
	End Property

	Public Property Get FgColor()
		FgColor = pHtmlDoc.fgColor
	End Property

	Public Property Get FileCreatedDate()
		FileCreatedDate = pHtmlDoc.fileCreatedDate
	End Property

	Public Property Get FileModifiedDate()
		FileModifiedDate = pHtmlDoc.fileModifiedDate
	End Property

	Public Property Get FileSize()
		FileSize = pHtmlDoc.fileSize
	End Property

	Public Property Get FileUpdatedDate()
		FileUpdatedDate = pHtmlDoc.fileUpdatedDate
	End Property

	Public Property Get FirstChild()
		Set FirstChild = pHtmlDoc.firstChild
	End Property

	Public Property Get Forms()
		Set Forms = pHtmlDoc.forms
	End Property

	Public Property Get Frames()
		Set Frames = pHtmlDoc.frames
	End Property

	Public Property Get Head()
		Set Head = pHtmlDoc.getElementsByTagName("head")(0)
	End Property

	Public Property Get HTML()
		HTML = pHtmlDoc.documentElement.outerHTML
	End Property

	Public Property Get Images()
		Set Images = pHtmlDoc.images
	End Property

	Public Property Get Implementation()
		Set Implementation = pHtmlDoc.implementation
	End Property

	Public Property Get LastChild()
		Set LastChild = pHtmlDoc.lastChild
	End Property

	Public Property Get LastModified()
		LastModified = pHtmlDoc.lastModified
	End Property

	Public Property Get LinkColor()
		Set LinkColor = pHtmlDoc.linkColor
	End Property

	Public Property Get Links()
		Set Links = pHtmlDoc.links
	End Property

	Public Property Get Location()
		Set Location = pHtmlDoc.location
	End Property

	Public Property Get Media()
		Media = pHtmlDoc.media
	End Property

	Public Property Get MimeType()
		On Error Resume Next
		Dim strMimeType
		strMimeType = pHtmlDoc.mimeType
		If Err.Number = 0 Then MimeType = strMimeType
	End Property

	Public Property Get Name()
		Name = pHtmlDoc.nameProp
	End Property
	
	Public Property Get Namespaces()
		Set Namespaces = pHtmlDoc.namespaces
	End Property

	Public Property Get NextSibling()
		Set NextSibling = pHtmlDoc.nextSibling
	End Property

	Public Property Get NodeName()
		NodeName = pHtmlDoc.nodeName
	End Property

	Public Property Get NodeType()
		NodeType = pHtmlDoc.nodeType
	End Property

	Public Property Get NodeValue()
		If IsObject(pHtmlDoc.nodeValue) Then
			Set NodeValue = pHtmlDoc.nodeValue
		Else
			NodeValue = pHtmlDoc.nodeValue
		End If
	End Property

	Public Property Get OwnerDocument()
		Set OwnerDocument = pHtmlDoc.ownerDocument
	End Property

	Public Property Get ParentNode()
		Set ParentNode = pHtmlDoc.parentNode
	End Property

	Public Property Get ParentWindow()
		Set ParentWindow = pHtmlDoc.parentWindow
	End Property

	Public Property Get Plugins()
		Set Plugins = pHtmlDoc.plugins
	End Property

	Public Property Get PreviousSibling()
		Set PreviousSibling = pHtmlDoc.previousSibling
	End Property

	Public Property Get ReadyState()
		ReadyState = pHtmlDoc.readyState
	End Property

	Public Property Get Referrer()
		Referrer = pHtmlDoc.referrer
	End Property

	Public Property Get Scripts()
		Set Scripts = pHtmlDoc.Scripts
	End Property

	Public Property Get Security()
		Security = pHtmlDoc.security
	End Property

	Public Property Get Selection()
		Set Selection = pHtmlDoc.selection
	End Property

	Public Property Get StyleSheets()
		Set StyleSheets = pHtmlDoc.styleSheets
	End Property

	Public Property Get Text()
		Text = pHtmlDoc.documentElement.innerText
	End Property

	Public Property Get Title()
		Title = pHtmlDoc.title
	End Property

	Public Property Get URL()
		URL = pHtmlDoc.url
	End Property

	Public Property Get URLUnencoded()
		URLUnencoded = pHtmlDoc.URLUnencoded
	End Property

	Public Property Get VLinkColor()
		Set VLinkColor = pHtmlDoc.vlinkColor
	End Property


	' Event Handlers


	Public Property Get OnActivate()
		Set OnActivate = pHtmlDoc.onactivate
	End Property

	Public Property Get OnAfterUpdate()
		Set OnAfterUpdate = pHtmlDoc.onafterupdate
	End Property

	Public Property Get OnBeforeActivate()
		Set OnBeforeActivate = pHtmlDoc.onbeforeactivate
	End Property

	Public Property Get OnBeforeDeactivate()
		Set OnBeforeDeactivate = pHtmlDoc.onbeforedeactivate
	End Property

	Public Property Get OnBeforeEditFocus()
		Set OnBeforeEditFocus = pHtmlDoc.onbeforeeditfocus
	End Property
	
	Public Property Get OnBeforeUpdate()
		Set OnBeforeUpdate = pHtmlDoc.onbeforeupdate
	End Property

	Public Property Get OnCellChange()
		Set OnCellChange = pHtmlDoc.oncellchange
	End Property

	Public Property Get OnClick()
		Set OnClick = pHtmlDoc.onclick
	End Property

	Public Property Get OnContextMenu()
		Set OnContextMenu = pHtmlDoc.oncontextmenu
	End Property

	Public Property Get OnControlSelect()
		Set OnControlSelect = pHtmlDoc.oncontrolselect
	End Property

	Public Property Get OnDataAvailable()
		Set OnDataAvailable = pHtmlDoc.ondataavailable
	End Property

	Public Property Get OnDatasetChanged()
		Set OnDatasetChanged = pHtmlDoc.ondatasetchanged
	End Property

	Public Property Get OnDatasetComplete()
		Set OnDatasetComplete = pHtmlDoc.ondatasetcomplete
	End Property

	Public Property Get OnDblClick()
		Set OnDblClick = pHtmlDoc.ondblclick
	End Property

	Public Property Get OnDeactivate()
		Set OnDeactivate = pHtmlDoc.ondeactivate
	End Property

	Public Property Get OnDragStart()
		Set OnDragStart = pHtmlDoc.ondragstart
	End Property

	Public Property Get onErrorUpdate()
		Set onErrorUpdate = pHtmlDoc.onerrorupdate
	End Property

	Public Property Get OnFocusIn()
		Set OnFocusIn = pHtmlDoc.onfocusin
	End Property

	Public Property Get OnFocusOut()
		Set OnFocusOut = pHtmlDoc.onfocusout
	End Property

	Public Property Get OnHelp()
		Set OnHelp = pHtmlDoc.onhelp
	End Property

	Public Property Get OnKeyDown()
		Set OnKeyDown = pHtmlDoc.onkeydown
	End Property

	Public Property Get OnKeyPress()
		Set OnKeyPress = pHtmlDoc.onkeypress
	End Property

	Public Property Get OnKeyUp()
		Set OnKeyUp = pHtmlDoc.onkeyup
	End Property

	Public Property Get OnMouseDown()
		Set OnMouseDown = pHtmlDoc.onmousedown
	End Property

	Public Property Get OnMouseMove()
		Set OnMouseMove = pHtmlDoc.onmousemove
	End Property

	Public Property Get OnMouseOut()
		Set OnMouseOut = pHtmlDoc.onmouseout
	End Property

	Public Property Get OnMouseOver()
		Set OnMouseOver = pHtmlDoc.onmouseover
	End Property

	Public Property Get OnMouseUp()
		Set OnMouseUp = pHtmlDoc.onmouseup
	End Property

	Public Property Get OnMouseWheel()
		Set OnMouseWheel = pHtmlDoc.onmousewheel
	End Property	

	Public Property Get OnMsSiteModeJumplistItemRemoved()
		Set OnMsSiteModeJumplistItemRemoved = pHtmlDoc.onmssitemodejumplistitemremoved
	End Property

	Public Property Get OnMsThumbnailClick()
		Set OnMsThumbnailClick = pHtmlDoc.onmsthumbnailclick
	End Property

	Public Property Get OnPropertyChange()
		Set OnPropertyChange = pHtmlDoc.onpropertychange
	End Property

	Public Property Get OnReadyStateChange()
		Set OnReadyStateChange = pHtmlDoc.onreadystatechange
	End Property

	Public Property Get OnRowEnter()
		Set OnRowEnter = pHtmlDoc.onrowenter
	End Property

	Public Property Get OnRowExit()
		Set OnRowExit = pHtmlDoc.onrowexit
	End Property

	Public Property Get OnRowsDelete()
		Set OnRowsDelete = pHtmlDoc.onrowsdelete
	End Property

	Public Property Get OnRowsInserted()
		Set OnRowsInserted = pHtmlDoc.onrowsinserted
	End Property

	Public Property Get OnSelectStart()
		Set OnSelectStart = pHtmlDoc.onselectstart
	End Property

	Public Property Get OnSelectionChange()
		Set OnSelectionChange = pHtmlDoc.onselectionchange
	End Property

	Public Property Get OnStop()
		Set OnStop = pHtmlDoc.onstop
	End Property

	Public Property Get OnStorage()
		Set OnStorage = pHtmlDoc.onstorage
	End Property

	Public Property Get OnStorageCommit()
		Set OnStorageCommit = pHtmlDoc.onstoragecommit
	End Property

	
	' Methods
	

	Public Function AdoptNode(objNode)
		Dim objNewNode
		Set objNewNode = ImportNode(objNode, True)
		objNode.parentElement.removeChild objNode
		Set AdoptNode = objNewNode
	End Function
	
	Public Function AppendChild(objChild)
		Set AppendChild = pHtmlDoc.appendChild(objChild)
	End Function
	
	Public Function AttachEvent(strEvent, objCallbackFunction)
		Set AttachEvent = pHtmlDoc.attachEvent(strEvent, objCallbackFunction)
	End Function
	
	Public Sub Clear()
		pHtmlDoc.clear()
	End Sub
	
	Public Function CloneNode(blnDeep)
		Set CloneNode = pHtmlDoc.cloneNode(blnDeep)
	End Function
	
	Public Function Close()
		pHtmlDoc.close()
	End Function
	
	Public Function CreateAttribute(strAttrName)
		Set CreateAttribute = pHtmlDoc.createAttribute(strAttrName)
	End Function

	Public Function CreateComment(strComment)
		Set CreateComment = pHtmlDoc.createComment(strComment)
	End Function
	
	Public Function CreateDocument()
		Set CreateDocument = CreateObject("HTMLFile")
	End Function
	
	Public Function CreateDocumentFragment()
		Set CreateDocumentFragment = pHtmlDoc.createDocumentFragment()
	End Function
	
	Public Function CreateDocumentFromURL(strURL)
		Dim objHtmlDoc
		Set objHtmlDoc = CreateObject("HTMLFile")
		objHtmlDoc.open strURL
		Set CreateDocumentFromURL = objHtmlDoc
	End Function
	
	Public Function CreateElement(strTag)
		Set CreateElement = pHtmlDoc.createElement(strTag)
	End Function
	
	Public Function CreateEvent()
		Set CreateEvent = pHtmlDoc.createEventObject()
	End Function
	
	Public Function CreateStyleSheet()
		Set CreateStyleSheet = pHtmlDoc.createStyleSheet()
	End Function
	
	Public Function CreateTextNode(strText)
		Set CreateTextNode = pHtmlDoc.createTextNode(strText)
	End Function
	
	Public Function DetachEvent(strEvent, objCallbackFunction)
		Set DetachEvent = pHtmlDoc.detachEvent(strEvent, objCallbackFunction)
	End Function
	
	Public Function ElementFromPoint(intX, intY)
		Set ElementFromPoint = pHtmlDoc.elementsFromPoint(intX, intY)(0)
	End Function
	
	Public Function ElementsFromPoint(intX, intY)
		Set ElementsFromPoint = pHtmlDoc.elementsFromPoint(intX, intY)
	End Function
	
	Public Function ExecCommand(strCmdID)
		Set ExecCommand = pHtmlDoc.execCommand(strCmdID)
	End Function

	Public Function ExecCommandShowHelp(strCmdID)
		Set ExecCommandShowHelp = pHtmlDoc.execCommandShowHelp(strCmdID)
	End Function
	
	Public Function ExecScript(strCode)
		Set ExecScript = pHtmlDoc.parentWindow.execScript(strCode)
	End Function
	
	Public Function FireEvent(strEventName)
		Set FireEvent = pHtmlDoc.fireEvent(strEventName)
	End Function
	
	Public Sub Focus()
		pHtmlDoc.focus()
	End Sub
	
	Public Function GetElementByID(strID)
		Set GetElementByID = pHtmlDoc.getElementById(strID)
	End Function
	
	Public Function GetElementsByClassName(strClassName)
		Dim objResultSet, _
			i

		Set objResultSet = New v_Data_Array

		For i = 0 To pHtmlDoc.all.length - 1
			If pHtmlDoc.all(i).className = strClassName Then
				objResultSet.Append pHtmlDoc.all(i)
			End If
		Next

		Set GetElementsByClassName = objResultSet
	End Function
	
	Public Function GetElementsByName(strName)
		Dim objResultSet, _
			i

		Set objResultSet = New v_Data_Array

		For i = 0 To pHtmlDoc.all.length - 1
			If pHtmlDoc.all(i).getAttribute("name") = strName Then
				objResultSet.Append pHtmlDoc.all(i)
			End If
		Next

		Set GetElementsByName = objResultSet
	End Function
	
	Public Function GetElementsByTagName(strTagName)
		Dim objResultSet, _
			objTagElements, _
			i

		Set objResultSet = New v_Data_Array
		Set objTagElements = pHtmlDoc.getElementsByTagName(strTagName)

		For i = 0 To objTagElements.Length - 1
			objResultSet.Append objTagElements(i)
		Next

		Set GetElementsByTagName = objResultSet
	End Function
	
	Public Function HasChildNodes()
		HasChildNodes = pHtmlDoc.hasChildNodes()
	End Function

	Public Function HasFocus()
		HasFocus = pHtmlDoc.hasFocus()
	End Function
	
	Public Function ImportNode(objNode, blnDeep)
		Select Case objNode.nodeType
			Case 1:
				Dim objNewNode, _
					i

				Set objNewNode = pHtmlDoc.createElement(objNode.nodeName)

				If Not objNode.attributes Is Nothing Then
					If objNode.attributes.length > 0 Then
						For i = 0 To objNode.attributes.length - 1
							If Not IsNull(objNode.getAttribute(objNode.attributes(i).nodeName)) And objNode.getAttribute(objNode.attributes(i).nodeName) <> "" Then 
								objNewNode.setAttribute objNode.attributes(i).nodeName, objNode.getAttribute(objNode.attributes(i).nodeName)
							End If
						Next
					End If
				End If

				If blnDeep And objNode.childNodes.length > 0 Then
					For i = 0 To objNode.childNodes.length - 1
						objNewNode.appendChild ImportNode(objNode.childNodes(i), True)
					Next
				End If
	
				Set ImportNode = objNewNode
			Case 3:
				Set ImportNode = pHtmlDoc.createTextNode(objNode.nodeValue)
			Case 8:
				Set ImportNode = pHtmlDoc.createComment(objNode.nodeValue)
		End Select	
	End Function

	Public Function InsertBefore(objNewChild, objRefChild)
		Set InsertBefore = pHtmlDoc.insertBefore(objNewChild, objRefChild)
	End Function

	Public Function IsEqualNode(objNode)
		Dim i

		If Not pHtmlDoc.nodeType = objNode.nodeType Then
			IsEqualNode = False
			Exit Function
		End If

		If Not pHtmlDoc.nodeName = objNode.nodeName Then
			IsEqualNode = False
			Exit Function
		End If

		If IsObject(pHtmlDoc.nodeValue) And IsObject(objNode.nodeValue) Then
			If Not pHtmlDoc.nodeValue Is objNode.nodeValue Then
				IsEqualNode = False
				Exit Function
			End If
		ElseIf Not IsObject(pHtmlDoc.nodeValue) And Not IsObject(objNode.nodeValue) Then
			If Not pHtmlDoc.nodeValue = objNode.nodeValue Then
				IsEqualNode = False
				Exit Function
			End If
		Else
			IsEqualNode = False
			Exit Function
		End If

		If Not pHtmlDoc.childNodes.length = objNode.childNodes.length Then
			IsEqualNode = False
			Exit Function
		Else
			For i = 0 To pHtmlDoc.childNodes.length - 1
				If Not pHtmlDoc.childNodes(i) Is objNode.childNodes(i) Then
					IsEqualNode = False
					Exit Function				
				End If
			Next
		End If

		If Not pHtmlDoc.attributes Is Nothing And Not objNode.attributes Is Nothing Then
			If Not pHtmlDoc.attributes.length = objNode.attributes.length Then
				IsEqualNode = False
				Exit Function
			Else
				Dim objAttr, _
					blnFound, _
					j

				For i = 0 To pHtmlDoc.attributes.length - 1
					Set objAttr = pHtmlDoc.attributes(i)
					blnFound = False

					For j = 0 To objNode.attributes.length - 1
						If objAttr.name = objNode.attributes(j).name And objNode.value = objNode.attributes(j).value Then
							blnFound = True
							Exit For
						End If		
					Next

					If Not blnFound Then 
						IsEqualNode = False
						Exit Function
					End If
				Next
			End If
		ElseIf Not (pHtmlDoc.attributes Is Nothing And objNode.attributes Is Nothing) Then
			IsEqualNode = False
			Exit Function	
		End If

		IsEqualNode = True
	End Function
	
	Public Function IsSameNode(objNode)
		If objNode Is pHtmlDoc Then
			IsEqualNode = True
		Else
			IsEqualNode = False
		End If
	End Function
	
	Public Function IsSupported(strFeature, strVersion)
		If strVersion = "1.0" Then
			IsSupported = pHtmlDoc.implementation.hasFeature(strFeature, "1.0")
		ElseIf strVersion = "2.0" Then
			IsSupported = pHtmlDoc.implementation.hasFeature(strFeature, "2.0")
		End If
	End Function

	Public Sub Load(strFile)
		Dim objFSO
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		WriteLn objFSO.OpenTextFile(strFile, 1).ReadAll()
		Set FSO = Nothing
	End Sub

	Public Sub Normalize()
		Dim objNode, _
			objNextNode

		Set objNode = pHtmlDoc.firstChild

		Do
			If objNode.nodeType = 3 Then
				Do
					If Not objNode.nextSibling Is Nothing Then
						Set objNextNode = objNode.nextSibling

						If objNextNode.nodeType = 3 Then					
							objNode.appendData objNextNode.data
							pHtmlDoc.removeChild objNextNode
						End If
					Else
						Set objNextNode = Nothing
					End If
				Loop While Not objNextNode Is Nothing
			End If
		
			If Not objNode.nextSibling Is Nothing Then
				Set objNode = objNode.nextSibling
			Else
				Set objNode = Nothing
			End If
		Loop While Not objNode Is Nothing
	End Sub
	
	Public Function QueryCommandEnabled(strCmdID)
		Set QueryCommandEnabled = pHtmlDoc.queryCommandEnabled(strCmdID)
	End Function
	
	Public Function QueryCommandIndeterm(strCmdID)
		Set QueryCommandIndeterm = pHtmlDoc.queryCommandIndeterm(strCmdID)
	End Function
	
	Public Function QueryCommandState(strCmdID)
		Set QueryCommandState = pHtmlDoc.queryCommandState(strCmdID)
	End Function
	
	Public Function QueryCommandSupported(strCmdID)
		Set QueryCommandSupported = pHtmlDoc.queryCommandSupported(strCmdID)
	End Function
	
	Public Function QueryCommandText(strCmdID)
		Set QueryCommandText = pHtmlDoc.queryCommandText(strCmdID)
	End Function
	
	Public Function QueryCommandValue(strCmdID)
		Set QueryCommandValue = pHtmlDoc.queryCommandValue(strCmdID)
	End Function
	
	Public Function QuerySelector(strSelector)
		Dim objStyleSheet, _
			objResult, _
			i

		Set objStyleSheet = pHtmlDoc.createStyleSheet()

		objStyleSheet.addRule strSelector, "k:v" 

		For i = 0 To pHtmlDoc.all.Length - 1
			If pHtmlDoc.all(i).currentStyle.getAttribute("k") = "v" Then
				Set objResult = pHtmlDoc.all(i)
				Exit For
			End If
		Next

		objStyleSheet.removeRule 0

		Set QuerySelector = objResult
	End Function
	
	Public Function QuerySelectorAll(strSelectors)
		Dim objStyleSheet, _
			objResultSet, _
			i, _
			j

		Set objStyleSheet = pHtmlDoc.createStyleSheet()
		Set objResultSet = New v_Data_Array

		strSelectors = Split(strSelectors, ",")

		For i = 0 To UBound(strSelectors)
			objStyleSheet.addRule strSelectors(i), "k:v" 

			For j = 0 To pHtmlDoc.all.length - 1
				If pHtmlDoc.all(j).currentStyle.getAttribute("k") = "v" Then
					objResultSet.Append pHtmlDoc.all(j)
				End If
			Next

			objStyleSheet.removeRule 0
		Next

		Set QuerySelectorAll = objResultSet
	End Function

	Public Sub Recalc()
		pHtmlDoc.recalc()
	End Sub

	Public Sub ReleaseCapture()
		pHtmlDoc.releaseCapture()
	End Sub
	
	Public Function RemoveChild(objChild)
		Set RemoveChild = pHtmlDoc.removeChild(objChild)
	End Function
	
	Public Function RemoveNode(blnDeep)
		Set RemoveNode = pHtmlDoc.removeNode(blnDeep)
	End Function

	Public Function ReplaceChild(objNewChild, objOldChild)
		Set ReplaceChild = pHtmlDoc.replaceChild(objNewChild, objOldChild)
	End Function
	
	Public Function ReplaceNode(objReplaceNode)
		Set ReplaceNode = pHtmlDoc.replaceNode(objReplaceNode)
	End Function
	
	Public Sub Save(strFilename)
		Dim objFSO, _
			objHtmlFile

		Set objFSO = CreateObject("Scripting.FileSystemObject")

		If objFSO.FileExists(strFilename) Then
			Set objHtmlFile = objFSO.OpenTextFile(strFilename, 2, True)			
		Else
			Set objHtmlFile = objFSO.CreateTextFile(strFilename, True)
		End If

		With objHtmlFile
			.WriteLine Me.HTML()
			.Close()
		End With
	End Sub
	
	Public Function SwapNode(objOtherNode)
		Set SwapNode = pHtmlDoc.swapNode(objOtherNode)
	End Function
	
	Public Function ToDocument()
		Set ToDocument = pHtmlDoc
	End Function

	Public Function ToString()
		ToString = pHtmlDoc.toString()
	End Function
	
	Public Sub UpdateSettings()
		pHtmlDoc.updateSettings()
	End Sub

	Public Sub Write(strText)
		pHtmlDoc.write strText
	End Sub
	
	Public Sub WriteLn(strText)
		pHtmlDoc.writeln strText
	End Sub


	' Destructor


	Private Sub Class_Terminate()
		Set pHtmlDoc = Nothing
	End Sub
End Class

If WScript.ScriptName = "v_HTML_Document.vbs" Then

End If
