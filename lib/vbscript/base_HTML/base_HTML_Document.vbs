Option Explicit

Include "base_Data_Array"

Class base_HTML_Document
	Private p_HtmlDoc


	' Constructor


	Private Sub Class_Initialize()
		Set p_HtmlDoc = CreateObject("HTMLFile")
	End Sub


	' Properties


	Public Property Get ActiveElement()
		Set ActiveElement = p_HtmlDoc.activeElement
	End Property

	Public Property Get ALinkColor()
		Set ALinkColor = p_HtmlDoc.alinkColor
	End Property

	Public Property Get All()
		Set All = p_HtmlDoc.all
	End Property

	Public Property Get Anchors()
		Set Anchors = p_HtmlDoc.anchors
	End Property

	Public Property Get Applets()
		Set Applets = p_HtmlDoc.applets
	End Property

	Public Property Get Attributes()
		Set Attributes = p_HtmlDoc.attributes
	End Property

	Public Property Get BgColor()
		BgColor = p_HtmlDoc.bgColor
	End Property

	Public Property Get Body()
		Set Body = p_HtmlDoc.body
	End Property

	Public Property Get CharacterSet()
		CharacterSet = p_HtmlDoc.charset
	End Property

	Public Property Get ChildNodes()
		Set ChildNodes = p_HtmlDoc.childNodes
	End Property

	Public Property Get Compatible()
		Set Compatible = p_HtmlDoc.compatible
	End Property

	Public Property Get CompatMode()
		CompatMode = p_HtmlDoc.compatMode
	End Property

	Public Property Get ContentType()
		On Error Resume Next
		Dim strContentType
		strContentType = p_HtmlDoc.mimeType
		If Err.Number = 0 Then ContentType = strContentType
	End Property

	Public Property Get Cookie()
		Cookie = p_HtmlDoc.cookie
	End Property

	Public Property Get DefaultCharset()
		DefaultCharset = p_HtmlDoc.defaultCharset
	End Property

	Public Property Get DefaultView()
		Set DefaultView = p_HtmlDoc.parentWindow
	End Property

	Public Property Get DesignMode()
		DesignMode = p_HtmlDoc.designMode
	End Property

	Public Property Get Dir()
		Dir = p_HtmlDoc.dir
	End Property

	Public Property Get DocType()
		Set DocType = p_HtmlDoc.doctype
	End Property

	Public Property Get DocumentElement()
		Set DocumentElement = p_HtmlDoc.documentElement
	End Property

	Public Property Get DocumentMode()
		Set DocumentMode = p_HtmlDoc.documentMode
	End Property

	Public Property Get DocumentURI()
		DocumentURI = p_HtmlDoc.url
	End Property

	Public Property Get Domain()
		Domain = p_HtmlDoc.domain
	End Property

	Public Property Get Embeds()
		Embeds = p_HtmlDoc.embeds
	End Property

	Public Property Get FgColor()
		FgColor = p_HtmlDoc.fgColor
	End Property

	Public Property Get FileCreatedDate()
		FileCreatedDate = p_HtmlDoc.fileCreatedDate
	End Property

	Public Property Get FileModifiedDate()
		FileModifiedDate = p_HtmlDoc.fileModifiedDate
	End Property

	Public Property Get FileSize()
		FileSize = p_HtmlDoc.fileSize
	End Property

	Public Property Get FileUpdatedDate()
		FileUpdatedDate = p_HtmlDoc.fileUpdatedDate
	End Property

	Public Property Get FirstChild()
		Set FirstChild = p_HtmlDoc.firstChild
	End Property

	Public Property Get Forms()
		Set Forms = p_HtmlDoc.forms
	End Property

	Public Property Get Frames()
		Set Frames = p_HtmlDoc.frames
	End Property

	Public Property Get Head()
		Set Head = p_HtmlDoc.getElementsByTagName("head")(0)
	End Property

	Public Property Get HTML()
		HTML = p_HtmlDoc.documentElement.outerHTML
	End Property

	Public Property Get Images()
		Set Images = p_HtmlDoc.images
	End Property

	Public Property Get Implementation()
		Set Implementation = p_HtmlDoc.implementation
	End Property

	Public Property Get LastChild()
		Set LastChild = p_HtmlDoc.lastChild
	End Property

	Public Property Get LastModified()
		LastModified = p_HtmlDoc.lastModified
	End Property

	Public Property Get LinkColor()
		Set LinkColor = p_HtmlDoc.linkColor
	End Property

	Public Property Get Links()
		Set Links = p_HtmlDoc.links
	End Property

	Public Property Get Location()
		Set Location = p_HtmlDoc.location
	End Property

	Public Property Get Media()
		Media = p_HtmlDoc.media
	End Property

	Public Property Get MimeType()
		On Error Resume Next
		Dim strMimeType
		strMimeType = p_HtmlDoc.mimeType
		If Err.Number = 0 Then MimeType = strMimeType
	End Property

	Public Property Get Name()
		Name = p_HtmlDoc.nameProp
	End Property
	
	Public Property Get Namespaces()
		Set Namespaces = p_HtmlDoc.namespaces
	End Property

	Public Property Get NextSibling()
		Set NextSibling = p_HtmlDoc.nextSibling
	End Property

	Public Property Get NodeName()
		NodeName = p_HtmlDoc.nodeName
	End Property

	Public Property Get NodeType()
		NodeType = p_HtmlDoc.nodeType
	End Property

	Public Property Get NodeValue()
		If IsObject(p_HtmlDoc.nodeValue) Then
			Set NodeValue = p_HtmlDoc.nodeValue
		Else
			NodeValue = p_HtmlDoc.nodeValue
		End If
	End Property

	Public Property Get OwnerDocument()
		Set OwnerDocument = p_HtmlDoc.ownerDocument
	End Property

	Public Property Get ParentNode()
		Set ParentNode = p_HtmlDoc.parentNode
	End Property

	Public Property Get ParentWindow()
		Set ParentWindow = p_HtmlDoc.parentWindow
	End Property

	Public Property Get Plugins()
		Set Plugins = p_HtmlDoc.plugins
	End Property

	Public Property Get PreviousSibling()
		Set PreviousSibling = p_HtmlDoc.previousSibling
	End Property

	Public Property Get ReadyState()
		ReadyState = p_HtmlDoc.readyState
	End Property

	Public Property Get Referrer()
		Referrer = p_HtmlDoc.referrer
	End Property

	Public Property Get Scripts()
		Set Scripts = p_HtmlDoc.Scripts
	End Property

	Public Property Get Security()
		Security = p_HtmlDoc.security
	End Property

	Public Property Get Selection()
		Set Selection = p_HtmlDoc.selection
	End Property

	Public Property Get StyleSheets()
		Set StyleSheets = p_HtmlDoc.styleSheets
	End Property

	Public Property Get Text()
		Text = p_HtmlDoc.documentElement.innerText
	End Property

	Public Property Get Title()
		Title = p_HtmlDoc.title
	End Property

	Public Property Get URL()
		URL = p_HtmlDoc.url
	End Property

	Public Property Get URLUnencoded()
		URLUnencoded = p_HtmlDoc.URLUnencoded
	End Property

	Public Property Get VLinkColor()
		Set VLinkColor = p_HtmlDoc.vlinkColor
	End Property


	' Event Handlers


	Public Property Get OnActivate()
		Set OnActivate = p_HtmlDoc.onactivate
	End Property

	Public Property Get OnAfterUpdate()
		Set OnAfterUpdate = p_HtmlDoc.onafterupdate
	End Property

	Public Property Get OnBeforeActivate()
		Set OnBeforeActivate = p_HtmlDoc.onbeforeactivate
	End Property

	Public Property Get OnBeforeDeactivate()
		Set OnBeforeDeactivate = p_HtmlDoc.onbeforedeactivate
	End Property

	Public Property Get OnBeforeEditFocus()
		Set OnBeforeEditFocus = p_HtmlDoc.onbeforeeditfocus
	End Property
	
	Public Property Get OnBeforeUpdate()
		Set OnBeforeUpdate = p_HtmlDoc.onbeforeupdate
	End Property

	Public Property Get OnCellChange()
		Set OnCellChange = p_HtmlDoc.oncellchange
	End Property

	Public Property Get OnClick()
		Set OnClick = p_HtmlDoc.onclick
	End Property

	Public Property Get OnContextMenu()
		Set OnContextMenu = p_HtmlDoc.oncontextmenu
	End Property

	Public Property Get OnControlSelect()
		Set OnControlSelect = p_HtmlDoc.oncontrolselect
	End Property

	Public Property Get OnDataAvailable()
		Set OnDataAvailable = p_HtmlDoc.ondataavailable
	End Property

	Public Property Get OnDatasetChanged()
		Set OnDatasetChanged = p_HtmlDoc.ondatasetchanged
	End Property

	Public Property Get OnDatasetComplete()
		Set OnDatasetComplete = p_HtmlDoc.ondatasetcomplete
	End Property

	Public Property Get OnDblClick()
		Set OnDblClick = p_HtmlDoc.ondblclick
	End Property

	Public Property Get OnDeactivate()
		Set OnDeactivate = p_HtmlDoc.ondeactivate
	End Property

	Public Property Get OnDragStart()
		Set OnDragStart = p_HtmlDoc.ondragstart
	End Property

	Public Property Get onErrorUpdate()
		Set onErrorUpdate = p_HtmlDoc.onerrorupdate
	End Property

	Public Property Get OnFocusIn()
		Set OnFocusIn = p_HtmlDoc.onfocusin
	End Property

	Public Property Get OnFocusOut()
		Set OnFocusOut = p_HtmlDoc.onfocusout
	End Property

	Public Property Get OnHelp()
		Set OnHelp = p_HtmlDoc.onhelp
	End Property

	Public Property Get OnKeyDown()
		Set OnKeyDown = p_HtmlDoc.onkeydown
	End Property

	Public Property Get OnKeyPress()
		Set OnKeyPress = p_HtmlDoc.onkeypress
	End Property

	Public Property Get OnKeyUp()
		Set OnKeyUp = p_HtmlDoc.onkeyup
	End Property

	Public Property Get OnMouseDown()
		Set OnMouseDown = p_HtmlDoc.onmousedown
	End Property

	Public Property Get OnMouseMove()
		Set OnMouseMove = p_HtmlDoc.onmousemove
	End Property

	Public Property Get OnMouseOut()
		Set OnMouseOut = p_HtmlDoc.onmouseout
	End Property

	Public Property Get OnMouseOver()
		Set OnMouseOver = p_HtmlDoc.onmouseover
	End Property

	Public Property Get OnMouseUp()
		Set OnMouseUp = p_HtmlDoc.onmouseup
	End Property

	Public Property Get OnMouseWheel()
		Set OnMouseWheel = p_HtmlDoc.onmousewheel
	End Property	

	Public Property Get OnMsSiteModeJumplistItemRemoved()
		Set OnMsSiteModeJumplistItemRemoved = p_HtmlDoc.onmssitemodejumplistitemremoved
	End Property

	Public Property Get OnMsThumbnailClick()
		Set OnMsThumbnailClick = p_HtmlDoc.onmsthumbnailclick
	End Property

	Public Property Get OnPropertyChange()
		Set OnPropertyChange = p_HtmlDoc.onpropertychange
	End Property

	Public Property Get OnReadyStateChange()
		Set OnReadyStateChange = p_HtmlDoc.onreadystatechange
	End Property

	Public Property Get OnRowEnter()
		Set OnRowEnter = p_HtmlDoc.onrowenter
	End Property

	Public Property Get OnRowExit()
		Set OnRowExit = p_HtmlDoc.onrowexit
	End Property

	Public Property Get OnRowsDelete()
		Set OnRowsDelete = p_HtmlDoc.onrowsdelete
	End Property

	Public Property Get OnRowsInserted()
		Set OnRowsInserted = p_HtmlDoc.onrowsinserted
	End Property

	Public Property Get OnSelectStart()
		Set OnSelectStart = p_HtmlDoc.onselectstart
	End Property

	Public Property Get OnSelectionChange()
		Set OnSelectionChange = p_HtmlDoc.onselectionchange
	End Property

	Public Property Get OnStop()
		Set OnStop = p_HtmlDoc.onstop
	End Property

	Public Property Get OnStorage()
		Set OnStorage = p_HtmlDoc.onstorage
	End Property

	Public Property Get OnStorageCommit()
		Set OnStorageCommit = p_HtmlDoc.onstoragecommit
	End Property

	
	' Methods
	

	Public Function AdoptNode(objNode)
		Dim objNewNode
		Set objNewNode = ImportNode(objNode, True)
		objNode.parentElement.removeChild objNode
		Set AdoptNode = objNewNode
	End Function
	
	Public Function AppendChild(objChild)
		Set AppendChild = p_HtmlDoc.appendChild(objChild)
	End Function
	
	Public Function AttachEvent(strEvent, objCallbackFunction)
		Set AttachEvent = p_HtmlDoc.attachEvent(strEvent, objCallbackFunction)
	End Function
	
	Public Sub Clear()
		p_HtmlDoc.clear()
	End Sub
	
	Public Function CloneNode(blnDeep)
		Set CloneNode = p_HtmlDoc.cloneNode(blnDeep)
	End Function
	
	Public Function Close()
		p_HtmlDoc.close()
	End Function
	
	Public Function CreateAttribute(strAttrName)
		Set CreateAttribute = p_HtmlDoc.createAttribute(strAttrName)
	End Function

	Public Function CreateComment(strComment)
		Set CreateComment = p_HtmlDoc.createComment(strComment)
	End Function
	
	Public Function CreateDocument()
		Set CreateDocument = CreateObject("HTMLFile")
	End Function
	
	Public Function CreateDocumentFragment()
		Set CreateDocumentFragment = p_HtmlDoc.createDocumentFragment()
	End Function
	
	Public Function CreateDocumentFromURL(strURL)
		Dim objHtmlDoc
		Set objHtmlDoc = CreateObject("HTMLFile")
		objHtmlDoc.open strURL
		Set CreateDocumentFromURL = objHtmlDoc
	End Function
	
	Public Function CreateElement(strTag)
		Set CreateElement = p_HtmlDoc.createElement(strTag)
	End Function
	
	Public Function CreateEvent()
		Set CreateEvent = p_HtmlDoc.createEventObject()
	End Function
	
	Public Function CreateStyleSheet()
		Set CreateStyleSheet = p_HtmlDoc.createStyleSheet()
	End Function
	
	Public Function CreateTextNode(strText)
		Set CreateTextNode = p_HtmlDoc.createTextNode(strText)
	End Function
	
	Public Function DetachEvent(strEvent, objCallbackFunction)
		Set DetachEvent = p_HtmlDoc.detachEvent(strEvent, objCallbackFunction)
	End Function
	
	Public Function ElementFromPoint(intX, intY)
		Set ElementFromPoint = p_HtmlDoc.elementsFromPoint(intX, intY)(0)
	End Function
	
	Public Function ElementsFromPoint(intX, intY)
		Set ElementsFromPoint = p_HtmlDoc.elementsFromPoint(intX, intY)
	End Function
	
	Public Function ExecCommand(strCmdID)
		Set ExecCommand = p_HtmlDoc.execCommand(strCmdID)
	End Function

	Public Function ExecCommandShowHelp(strCmdID)
		Set ExecCommandShowHelp = p_HtmlDoc.execCommandShowHelp(strCmdID)
	End Function
	
	Public Function ExecScript(strCode)
		Set ExecScript = p_HtmlDoc.parentWindow.execScript(strCode)
	End Function
	
	Public Function FireEvent(strEventName)
		Set FireEvent = p_HtmlDoc.fireEvent(strEventName)
	End Function
	
	Public Sub Focus()
		p_HtmlDoc.focus()
	End Sub
	
	Public Function GetElementByID(strID)
		Set GetElementByID = p_HtmlDoc.getElementById(strID)
	End Function
	
	Public Function GetElementsByClassName(strClassName)
		Dim objResultSet, _
			i

		Set objResultSet = New base_Data_Array

		For i = 0 To p_HtmlDoc.all.length - 1
			If p_HtmlDoc.all(i).className = strClassName Then
				objResultSet.Append p_HtmlDoc.all(i)
			End If
		Next

		Set GetElementsByClassName = objResultSet
	End Function
	
	Public Function GetElementsByName(strName)
		Dim objResultSet, _
			i

		Set objResultSet = New base_Data_Array

		For i = 0 To p_HtmlDoc.all.length - 1
			If p_HtmlDoc.all(i).getAttribute("name") = strName Then
				objResultSet.Append p_HtmlDoc.all(i)
			End If
		Next

		Set GetElementsByName = objResultSet
	End Function
	
	Public Function GetElementsByTagName(strTagName)
		Dim objResultSet, _
			objTagElements, _
			i

		Set objResultSet = New base_Data_Array
		Set objTagElements = p_HtmlDoc.getElementsByTagName(strTagName)

		For i = 0 To objTagElements.Length - 1
			objResultSet.Append objTagElements(i)
		Next

		Set GetElementsByTagName = objResultSet
	End Function
	
	Public Function HasChildNodes()
		HasChildNodes = p_HtmlDoc.hasChildNodes()
	End Function

	Public Function HasFocus()
		HasFocus = p_HtmlDoc.hasFocus()
	End Function
	
	Public Function ImportNode(objNode, blnDeep)
		Select Case objNode.nodeType
			Case 1:
				Dim objNewNode, _
					i

				Set objNewNode = p_HtmlDoc.createElement(objNode.nodeName)

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
				Set ImportNode = p_HtmlDoc.createTextNode(objNode.nodeValue)
			Case 8:
				Set ImportNode = p_HtmlDoc.createComment(objNode.nodeValue)
		End Select	
	End Function

	Public Function InsertBefore(objNewChild, objRefChild)
		Set InsertBefore = p_HtmlDoc.insertBefore(objNewChild, objRefChild)
	End Function

	Public Function IsEqualNode(objNode)
		Dim i

		If Not p_HtmlDoc.nodeType = objNode.nodeType Then
			IsEqualNode = False
			Exit Function
		End If

		If Not p_HtmlDoc.nodeName = objNode.nodeName Then
			IsEqualNode = False
			Exit Function
		End If

		If IsObject(p_HtmlDoc.nodeValue) And IsObject(objNode.nodeValue) Then
			If Not p_HtmlDoc.nodeValue Is objNode.nodeValue Then
				IsEqualNode = False
				Exit Function
			End If
		ElseIf Not IsObject(p_HtmlDoc.nodeValue) And Not IsObject(objNode.nodeValue) Then
			If Not p_HtmlDoc.nodeValue = objNode.nodeValue Then
				IsEqualNode = False
				Exit Function
			End If
		Else
			IsEqualNode = False
			Exit Function
		End If

		If Not p_HtmlDoc.childNodes.length = objNode.childNodes.length Then
			IsEqualNode = False
			Exit Function
		Else
			For i = 0 To p_HtmlDoc.childNodes.length - 1
				If Not p_HtmlDoc.childNodes(i) Is objNode.childNodes(i) Then
					IsEqualNode = False
					Exit Function				
				End If
			Next
		End If

		If Not p_HtmlDoc.attributes Is Nothing And Not objNode.attributes Is Nothing Then
			If Not p_HtmlDoc.attributes.length = objNode.attributes.length Then
				IsEqualNode = False
				Exit Function
			Else
				Dim objAttr, _
					blnFound, _
					j

				For i = 0 To p_HtmlDoc.attributes.length - 1
					Set objAttr = p_HtmlDoc.attributes(i)
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
		ElseIf Not (p_HtmlDoc.attributes Is Nothing And objNode.attributes Is Nothing) Then
			IsEqualNode = False
			Exit Function	
		End If

		IsEqualNode = True
	End Function
	
	Public Function IsSameNode(objNode)
		If objNode Is p_HtmlDoc Then
			IsEqualNode = True
		Else
			IsEqualNode = False
		End If
	End Function
	
	Public Function IsSupported(strFeature, strVersion)
		If strVersion = "1.0" Then
			IsSupported = p_HtmlDoc.implementation.hasFeature(strFeature, "1.0")
		ElseIf strVersion = "2.0" Then
			IsSupported = p_HtmlDoc.implementation.hasFeature(strFeature, "2.0")
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

		Set objNode = p_HtmlDoc.firstChild

		Do
			If objNode.nodeType = 3 Then
				Do
					If Not objNode.nextSibling Is Nothing Then
						Set objNextNode = objNode.nextSibling

						If objNextNode.nodeType = 3 Then					
							objNode.appendData objNextNode.data
							p_HtmlDoc.removeChild objNextNode
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
		Set QueryCommandEnabled = p_HtmlDoc.queryCommandEnabled(strCmdID)
	End Function
	
	Public Function QueryCommandIndeterm(strCmdID)
		Set QueryCommandIndeterm = p_HtmlDoc.queryCommandIndeterm(strCmdID)
	End Function
	
	Public Function QueryCommandState(strCmdID)
		Set QueryCommandState = p_HtmlDoc.queryCommandState(strCmdID)
	End Function
	
	Public Function QueryCommandSupported(strCmdID)
		Set QueryCommandSupported = p_HtmlDoc.queryCommandSupported(strCmdID)
	End Function
	
	Public Function QueryCommandText(strCmdID)
		Set QueryCommandText = p_HtmlDoc.queryCommandText(strCmdID)
	End Function
	
	Public Function QueryCommandValue(strCmdID)
		Set QueryCommandValue = p_HtmlDoc.queryCommandValue(strCmdID)
	End Function
	
	Public Function QuerySelector(strSelector)
		Dim objStyleSheet, _
			objResult, _
			i

		Set objStyleSheet = p_HtmlDoc.createStyleSheet()

		objStyleSheet.addRule strSelector, "k:v" 

		For i = 0 To p_HtmlDoc.all.Length - 1
			If p_HtmlDoc.all(i).currentStyle.getAttribute("k") = "v" Then
				Set objResult = p_HtmlDoc.all(i)
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

		Set objStyleSheet = p_HtmlDoc.createStyleSheet()
		Set objResultSet = New base_Data_Array

		strSelectors = Split(strSelectors, ",")

		For i = 0 To UBound(strSelectors)
			objStyleSheet.addRule strSelectors(i), "k:v" 

			For j = 0 To p_HtmlDoc.all.length - 1
				If p_HtmlDoc.all(j).currentStyle.getAttribute("k") = "v" Then
					objResultSet.Append p_HtmlDoc.all(j)
				End If
			Next

			objStyleSheet.removeRule 0
		Next

		Set QuerySelectorAll = objResultSet
	End Function

	Public Sub Recalc()
		p_HtmlDoc.recalc()
	End Sub

	Public Sub ReleaseCapture()
		p_HtmlDoc.releaseCapture()
	End Sub
	
	Public Function RemoveChild(objChild)
		Set RemoveChild = p_HtmlDoc.removeChild(objChild)
	End Function
	
	Public Function RemoveNode(blnDeep)
		Set RemoveNode = p_HtmlDoc.removeNode(blnDeep)
	End Function

	Public Function ReplaceChild(objNewChild, objOldChild)
		Set ReplaceChild = p_HtmlDoc.replaceChild(objNewChild, objOldChild)
	End Function
	
	Public Function ReplaceNode(objReplaceNode)
		Set ReplaceNode = p_HtmlDoc.replaceNode(objReplaceNode)
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
		Set SwapNode = p_HtmlDoc.swapNode(objOtherNode)
	End Function
	
	Public Function ToDocument()
		Set ToDocument = p_HtmlDoc
	End Function

	Public Function ToString()
		ToString = p_HtmlDoc.toString()
	End Function
	
	Public Sub UpdateSettings()
		p_HtmlDoc.updateSettings()
	End Sub

	Public Sub Write(strText)
		p_HtmlDoc.write strText
	End Sub
	
	Public Sub WriteLn(strText)
		p_HtmlDoc.writeln strText
	End Sub


	' Destructor


	Private Sub Class_Terminate()
		Set p_HtmlDoc = Nothing
	End Sub
End Class

If WScript.ScriptName = "v_HTML_Document.vbs" Then

End If
