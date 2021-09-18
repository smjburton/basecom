Option Explicit

Include "base_Data_Array"

Class base_HTML_Parser
	Private pHtmlParser, _
		pSelection, _
		pSanitize

	Private Sub Class_Initialize()
		Set pHtmlParser = CreateObject("HTMLFile")
		Set pSelection = New base_Data_Array
		pSanitize = True
	End Sub


	' Properties


	Public Property Get Tag()
		Tag = pSelection(0).tagName
	End Property

	Public Property Get Element(intIndex)
		Set Element = pSelection(intIndex)
	End Property

	Public Property Get Attribute(strAttr)
		If IsObject(pSelection(0).getAttribute(strAttr)) Then
			Set Attribute = pSelection(0).getAttribute(strAttr)
		Else
			Attribute = pSelection(0).getAttribute(strAttr)
		End If
	End Property

	Public Property Get Attributes()
		Set Attributes = pSelection(0).attributes
	End Property

	Public Property Get ID()
		ID = pSelection(0).id
	End Property

	Public Property Get ClassName()
		ClassName = pSelection(0).className
	End Property

	Public Property Get CurrentStyle()
		Set CurrentStyle = pSelection(0).currentStyle
	End Property

	Public Property Get Style()
		Set Style = pSelection(0).style
	End Property

	Public Property Get NodeName()
		NodeName = pSelection(0).nodeName
	End Property

	Public Property Get NodeType()
		NodeType = pSelection(0).nodeType
	End Property

	Public Property Get NodeValue()
		If IsObject(pSelection(0).nodeValue) Then
			Set NodeValue = pSelection(0).nodeValue
		Else
			NodeValue = pSelection(0).nodeValue
		End If
	End Property

	Public Default Property Get Item(intIndex)
		Set Item = HTMLParser(pSelection(intIndex))
	End Property

	Public Property Get Count()
		Count = pSelection.Length
	End Property

	Public Property Get Text()
		Dim strText, _
			i

		For i = 0 To pSelection.Length - 1
			strText = strText & pSelection(i).innerText
		Next

		Text = strText
	End Property

	Public Property Get HTML()
		Dim strHTML, _
			i

		For i = 0 To pSelection.Length - 1
			strHTML = strHTML & pSelection(i).outerHTML
		Next

		HTML = strHTML
	End Property

	Public Property Get Title()
		Title = pHtmlParser.Title
	End Property


	' Options


	' Option to sanitize the incoming HTML of any unsafe code.
	Public Property Get Sanitize()
		Sanitize = pSanitize
	End Property

	Public Property Let Sanitize(blnSanitize)
		pSanitize = blnSanitize
	End Property


	' Preconfigured Selections:


	Public Property Get A()
		Set A = ElementsByTagName("a")
	End Property

	Public Property Get Body()
		Set Body = ElementsByTagName("body")
	End Property

	Public Property Get Buttons()
		Set Buttons = ElementsByTagName("button")
	End Property

	Public Property Get Comments()
		Dim objChildElements, _
			objResultSet, _
			i, _
			j

		Set objResultSet = New base_Data_Array

		For i = 0 to pSelection.Length - 1
			Set objChildElements = pSelection(i).getElementsByTagName("*") 		

			For j = 0 To objChildElements.length - 1
				If objChildElements(j).nodeType = 8 Then objResultSet.Append objChildElements(j)
			Next
		Next

		Set Comments = HTMLParser(objResultSet)
	End Property

	Public Property Get Divs()
		Set Divs = ElementsByTagName("div")
	End Property

	Public Property Get Forms()
		Set Forms = ElementsByTagName("form")
	End Property

	Public Property Get H1()
		Set H1 = ElementsByTagName("h1")
	End Property

	Public Property Get H2()
		Set H2 = ElementsByTagName("h2")
	End Property

	Public Property Get H3()
		Set H3 = ElementsByTagName("h3")
	End Property

	Public Property Get H4()
		Set H4 = ElementsByTagName("h4")
	End Property

	Public Property Get H5()
		Set H5 = ElementsByTagName("h5")
	End Property

	Public Property Get H6()
		Set H6 = ElementsByTagName("h6")
	End Property

	Public Property Get Head()
		Set Head = ElementsByTagName("head")
	End Property

	Public Property Get Images()
		Set Images = ElementsByTagName("img")
	End Property

	Public Property Get Inputs()
		Set Inputs = ElementsByTagName("input")
	End Property

	Public Property Get Labels()
		Set Labels = ElementsByTagName("label")
	End Property

	Public Property Get LI()
		Set LI = ElementsByTagName("li")
	End Property

	Public Property Get Links()
		Set Links = ElementsByTagName("a")
	End Property

	Public Property Get Meta()
		Set Meta = ElementsByTagName("meta")
	End Property

	Public Property Get OL()
		Set OL = ElementsByTagName("ol")
	End Property

	Public Property Get P()
		Set P = ElementsByTagName("p")
	End Property

	' * Doesn't work if you have pSanitize = True
	Public Property Get Scripts()
		If Not pSanitize Then
			Set Scripts = ElementsByTagName("script")
		Else
			' Error
		End If		
	End Property

	Public Property Get Spans()
		Set Spans = ElementsByTagName("span")
	End Property

	Public Property Get Styles()
		Set Styles = ElementsByTagName("style")
	End Property

	' Retrieves all '<link>' and '<style>' elements.
	Public Property Get StyleSheets()
		' Set StyleSheets = pSelection.StyleSheets
	End Property

	Public Property Get Tables()
		Set Tables = ElementsByTagName("table")
	End Property

	Public Property Get UL()
		Set UL = ElementsByTagName("ul")
	End Property


	' Methods

	
	' Traversal Methods


	Public Function Ascend(intLevel, intIndex)
		Set Ascend = Offset(intLevel, intIndex)
	End Function

	Public Function Descend(intLevel, intIndex)
		Set Descend = Offset(-intLevel, intIndex)
	End Function

	Public Function PreviousNode(intIndex)
		Set PreviousNode = Offset(0, -intIndex)
	End Function

	Public Function NextNode(intIndex)
		Set NextNode = Offset(0, intIndex)
	End Function

	Public Function Offset(intLevel, intIndex)
		Dim strOffset, _
			i, _
			j

		For i = 0 To pHtmlParser.all.length - 1
			If pSelection(0) Is pHtmlParser.all(i) Then
				strOffset = "pHtmlParser.all(i)"

				If intLevel > 0 Then
					For j = 0 To intLevel - 1
						strOffset = strOffset & ".parentElement"
					Next
				ElseIf intLevel < 0 Then
					For j = 0 To intLevel + 1 Step -1
						strOffset = strOffset & ".firstChild"
					Next
				End If

				If intIndex > 0 Then
					For j = 0 To intIndex - 1
						strOffset = strOffset & ".nextSibling"
					Next
				ElseIf intIndex < 0 Then
					For j = 0 To intIndex + 1 Step -1
						strOffset = strOffset & ".previousSibling"
					Next
				End If

				Exit For
			End If
		Next

		Set Offset = HTMLParser(Eval(strOffset))
	End Function


	' Familial Methods


	Public Function Ancestor(intLevel, intIndex)
		Set Ancestor = Offset(intLevel, intIndex)
	End Function

	Public Function Descendant(intLevel, intIndex)
		Set Descendant = Offset(-intLevel, intIndex)
	End Function

	Public Function Parents(intLevel)
		Set Parents = Offset(intLevel, 0)
	End Function

	Public Function Child(intIndex)
		Set Child = HTMLParser(pSelection(0).Children(intIndex))
	End Function

	Public Function FirstChild()
		Set FirstChild = HTMLParser(pSelection(0).firstChild)
	End Function
	
	Public Function LastChild()
		Set LastChild = HTMLParser(pSelection(0).lastChild)
	End Function

	Public Function Sibling(intIndex)
		Set Sibling = Offset(0, intIndex)
	End Function

	Public Function Parent()
		Set Parent = HTMLParser(pSelection(0).parentNode)
	End Function

	Public Function Children()
		Set Children = HTMLParser(pSelection(0).Children)
	End Function

	Public Function Siblings()
		Set Siblings = HTMLParser(pSelection(0).parentNode.Children)
	End Function

	Public Function NextSibling()
		Set NextSibling = HTMLParser(pSelection(0).nextSibling)
	End Function

	Public Function PreviousSibling()
		Set PreviousSibling = HTMLParser(pSelection(0).previousSibling)
	End Function

	Public Function Relative(intLevel, intIndex)
		Set Relative = Offset(intLevel, intIndex)
	End Function


	' Selection Methods


	Public Property Get All()
		Set All = ElementsByTagName("*")
	End Property

	Public Function First()
		Set First = HTMLParser(pSelection(0))
	End Function

	Public Function Last()
		Set Last = HTMLParser(pSelection(pSelection.Length - 1))
	End Function

	Public Function UpTo(varCondition)
		Dim intEnd, _
			i

		If TypeName(varCondition) = "String" Then
			Dim objStyleSheet
			Set objStyleSheet = pHtmlParser.createStyleSheet()

			objStyleSheet.addRule varCondition, "k:v" 

			For i = 0 To pSelection.Length - 1
				If pSelection(i).currentStyle.getAttribute("k") = "v" Then
					intEnd = i
					Exit For
				End If
			Next

			objStyleSheet.removeRule 0
		ElseIf varCondition.nodeType = 1 Or varCondition.nodeType = 3 Or varCondition.nodeType = 8 Then
			For i = 0 To pSelection.Length - 1
				If varCondition Is pSelection(i) Then
					intEnd = i
					Exit For
				End If
			Next
		End If

		Set UpTo = Slice(0, intEnd)
	End Function

	Public Function Slice(intStart, intEnd)
		Set Slice = HTMLParser(pSelection.Slice(intStart, intEnd))
	End Function

	Public Function Has(strAttr)
		If Not IsNull(pSelection(0).getAttribute(strAttr)) Then
			Has = True
		Else
			Has = False
		End If
	End Function

	Public Function Contains(varCondition)
		Dim blnContains, _
			i, _
			j

		blnContains = False

		If TypeName(varCondition) = "String" Then
			Dim objStyleSheet, _
				objChildElements

			Set objStyleSheet = pHtmlParser.createStyleSheet()
			objStyleSheet.addRule varCondition, "k:v" 

			For i = 0 To pSelection.Length - 1
				Set objChildElements = pSelection(i).getElementsByTagName("*")

				For j = 0 To objChildElements.length - 1
					If objChildElements(j).currentStyle.getAttribute("k") = "v" Then
						blnContains = True
						Exit For
					End If
				Next
			Next

			objStyleSheet.removeRule 0
		ElseIf varCondition.nodeType = 1 Or varCondition.nodeType = 3 Or varCondition.nodeType = 8 Then
			For i = 0 To pSelection.Length - 1
				Set objChildElements = pSelection(i).getElementsByTagName("*")

				For j = 0 To objChildElements.length - 1
					If varCondition Is objChildElements(j) Then
						blnContains = True
						Exit For
					End If
				Next
			Next
		End If

		Contains = blnContains
	End Function

	Public Function IsSelected(varCondition)
		Dim blnSelected, _
			i		

		blnSelected = False

		If TypeName(varCondition) = "String" Then
			Dim objStyleSheet
			Set objStyleSheet = pHtmlParser.createStyleSheet()

			objStyleSheet.addRule varCondition, "k:v" 

			For i = 0 To pSelection.Length - 1
				If pSelection(i).currentStyle.getAttribute("k") = "v" Then
					blnSelected = True
					Exit For
				End If
			Next

			objStyleSheet.removeRule 0
		ElseIf varCondition.nodeType = 1 Or varCondition.nodeType = 3 Or varCondition.nodeType = 8 Then
			For i = 0 To pSelection.Length - 1
				If varCondition Is pSelection(i) Then
					blnSelected = True
					Exit For
				End If
			Next
		End If

		IsSelected = blnSelected
	End Function

	Public Function Exclude(varCondition)
		Dim objArray, _
			i

		Set objArray = New base_Data_Array

		If TypeName(varCondition) = "String" Then
			Dim objStyleSheet
			Set objStyleSheet = pHtmlParser.createStyleSheet()

			objStyleSheet.addRule varCondition, "k:v" 

			For i = 0 To pSelection.Length - 1
				If IsNull(pSelection(i).currentStyle.getAttribute("k")) Then
					objArray.Append pSelection(i)
				End If
			Next

			objStyleSheet.removeRule 0
		ElseIf varCondition.nodeType = 1 Or varCondition.nodeType = 3 Or varCondition.nodeType = 8 Then
			For i = 0 To pSelection.Length - 1
				If Not varCondition Is pSelection(i) Then
					objArray.Append pSelection(i)
				End If
			Next
		End If

		Set Exclude = HTMLParser(objArray)
	End Function
	
	Public Function Limit(intLimit)
		Set Limit = HTMLParser(pSelection.Slice(0, intLimit - 1))		
	End Function


	' Retrieval Methods


	Public Function ElementsByName(strName)
		Dim objResultSet, _
			objNameElements, _
			i, _
			j

		Set objResultSet = New base_Data_Array

		For i = 0 To pSelection.Length - 1
			Set objNameElements = pSelection(i).getElementsByTagName("*")

			For j = 0 To objNameElements.length - 1
				If objNameElements(j).getAttribute("name") = strName Then
					objResultSet.Append objNameElements(j)
				End If
			Next
		Next

		Set ElementsByName = HTMLParser(objResultSet)
	End Function

	Public Function ElementsByTagName(strTagName)
		Dim objResultSet, _
			objTagElements, _
			i, _
			j

		Set objResultSet = New base_Data_Array

		For i = 0 To pSelection.Length - 1
			Set objTagElements = pSelection(i).getElementsByTagName(strTagName)

			For j = 0 To objTagElements.length - 1
				objResultSet.Append objTagElements(j)
			Next
		Next

		Set ElementsByTagName = HTMLParser(objResultSet)
	End Function
	
	Public Function ElementsByClassName(strClassName)
		Dim objResultSet, _
			objClassElements, _
			i, _
			j

		Set objResultSet = New base_Data_Array

		For i = 0 To pSelection.Length - 1
			Set objClassElements = pSelection(i).getElementsByTagName("*")

			For j = 0 To objClassElements.length - 1
				If objClassElements(j).className = strClassName Then
					objResultSet.Append objClassElements(j)
				End If
			Next
		Next

		Set ElementsByClassName = HTMLParser(objResultSet)
	End Function

	Public Function ElementsByAttribute(strAttr)
		Dim objResultSet, _
			objAttrElements, _
			i, _
			j

		Set objResultSet = New base_Data_Array

		For i = 0 To pSelection.Length - 1
			Set objAttrElements = pSelection(i).getElementsByTagName("*")

			For j = 0 To objAttrElements.length - 1
				If Not IsNull(objAttrElements(j).getAttribute(strAttr)) Then
					If TypeName(objAttrElements(j).getAttribute(strAttr)) = "String" Then
						objResultSet.Append objAttrElements(j)
					End If
				End If
			Next
		Next

		Set ElementsByAttribute = HTMLParser(objResultSet)
	End Function

	Public Function ElementsByAttributeValue(strAttr, strValue)
		Dim objResultSet, _
			objAttrElements, _
			i, _
			j

		Set objResultSet = New base_Data_Array

		For i = 0 To pSelection.Length - 1
			Set objAttrElements = pSelection(i).getElementsByTagName("*")

			For j = 0 To objAttrElements.length - 1
				If Not IsNull(objAttrElements(j).getAttribute(strAttr)) Then
					If TypeName(objAttrElements(j).getAttribute(strAttr)) = "String" Then
						If objAttrElements(j).getAttribute(strAttr) = strValue Then
							objResultSet.Append objAttrElements(j)
						End If
					End If
				End If
			Next
		Next

		Set ElementsByAttributeValue = HTMLParser(objResultSet)
	End Function
	
	Public Function ElementByID(strID)
		Dim objResult, _
			objElements, _
			i, _
			j

		For i = 0 To pSelection.Length - 1
			Set objElements = pSelection(i).getElementsByTagName("*")

			For j = 0 To objElements.length - 1
				If objElements(j).id = strID Then
					Set objResult = objElements(j)
					Exit For
				End If
			Next
		Next

		Set ElementByID = HTMLParser(objResult)
	End Function

	Public Function Selector(strSelector)
		Dim objStyleSheet, _
			objChildElements, _
			objResult, _
			i, _
			j

		Set objStyleSheet = pHtmlParser.createStyleSheet()

		objStyleSheet.addRule strSelector, "k:v" 

		For i = 0 To pSelection.Length - 1
			Set objChildElements = pSelection(i).getElementsByTagName("*")

			For j = 0 To objChildElements.length - 1
				If objChildElements(j).currentStyle.getAttribute("k") = "v" Then
					Set objResult = objChildElements(j)
					Exit For
				End If
			Next
		Next

		objStyleSheet.removeRule 0

		Set Selector = HTMLParser(objResult)
	End Function

	Public Function SelectorAll(strSelectors)
		Dim objStyleSheet, _
			objChildElements, _
			objResultSet, _
			i, _
			j, _
			k

		Set objStyleSheet = pHtmlParser.createStyleSheet()
		Set objResultSet = New base_Data_Array

		strSelectors = Split(strSelectors, ",")

		For i = 0 To UBound(strSelectors)
			objStyleSheet.addRule strSelectors(i), "k:v" 

			For j = 0 To pSelection.Length - 1
				Set objChildElements = pSelection(j).getElementsByTagName("*")

				For k = 0 To objChildElements.length - 1
					If objChildElements(k).currentStyle.getAttribute("k") = "v" Then
						objResultSet.Append objChildElements(k)
					End If
				Next
			Next

			objStyleSheet.removeRule 0
		Next

		Set SelectorAll = HTMLParser(objResultSet)
	End Function

	Public Function Match(strType, strPattern, blnIgnoreCase, blnGlobal)
		Dim strInput, _
			intInputLimit, _
			objRegex, _
			objMatches, _
			objChildElements, _
			objResult, _
			i, _
			j, _
			k

		Set objRegex = New RegExp

		With objRegex
			.IgnoreCase = blnIgnoreCase
			.Global = blnGlobal
			.Pattern = strPattern
		End With

		For i = 0 To pSelection.Length - 1
			Set objChildElements = pSelection(i).getElementsByTagName("*")

			For j = 0 To objChildElements.length - 1
				If Not IsEmpty(objResult) Then Exit For

				If strType = "HTML" Then
					strInput = objChildElements(j).outerHTML
					intInputLimit = InStr(objChildElements(j).outerHTML, ">")
				ElseIf strType = "Text" Then
					If InStr(objChildElements(j).innerHTML, "<") > 0 Then
						strInput = Mid(objChildElements(j).innerHTML, 1, InStr(objChildElements(j).innerHTML, "<"))
						intInputLimit = InStr(objChildElements(j).innerHTML, "<")
					Else
						strInput = objChildElements(j).innerText
						intInputLimit = Len(objChildElements(j).innerText)
					End If
				ElseIf strType = "Element" Then
					strInput = Replace(objChildElements(j).outerHTML, objChildElements(j).innerHTML, "", 1, 1)
					intInputLimit = InStr(objChildElements(j).outerHTML, ">")
				ElseIf strType = "Attribute" Then
					Dim intStart, _
						intEnd

					intStart = InStr(objChildElements(j).outerHTML, "<") + Len(objChildElements(j).tagName) + 1
					intEnd = InStr(objChildElements(j).outerHTML, ">")

					intInputLimit = intEnd - intStart

					If Not intStart = intEnd Then
						strInput = Mid(objChildElements(j).outerHTML, intStart, intInputLimit)
					Else
						strInput = ""
					End If
				End If
				
				If objRegex.Test(strInput) Then
					Set objMatches = objRegex.Execute(strInput)

					For k = 0 To objMatches.Count - 1
						If objMatches(k).FirstIndex < intInputLimit Then
							Set objResult = objChildElements(j)
							Exit For
						End If
					Next
				End If
			Next
		Next

		Set Match = HTMLParser(objResult)
	End Function

	Public Function MatchAll(strType, strPattern, blnIgnoreCase, blnGlobal)
		Dim strInput, _
			intInputLimit, _
			objRegex, _
			objMatches, _
			objChildElements, _
			objResultSet, _
			i, _
			j, _
			k

		Set objRegex = New RegExp
		Set objResultSet = New base_Data_Array

		With objRegex
			.IgnoreCase = blnIgnoreCase
			.Global = blnGlobal
			.Pattern = strPattern
		End With

		For i = 0 To pSelection.Length - 1
			Set objChildElements = pSelection(i).getElementsByTagName("*")

			For j = 0 To objChildElements.length - 1
				If strType = "HTML" Then
					strInput = objChildElements(j).outerHTML
					intInputLimit = InStr(objChildElements(j).outerHTML, ">")
				ElseIf strType = "Text" Then
					If InStr(objChildElements(j).innerHTML, "<") > 0 Then
						strInput = Mid(objChildElements(j).innerHTML, 1, InStr(objChildElements(j).innerHTML, "<"))
						intInputLimit = InStr(objChildElements(j).innerHTML, "<")
					Else
						strInput = objChildElements(j).innerText
						intInputLimit = Len(objChildElements(j).innerText)
					End If
				ElseIf strType = "Element" Then
					strInput = Replace(objChildElements(j).outerHTML, objChildElements(j).innerHTML, "", 1, 1)
					intInputLimit = InStr(objChildElements(j).outerHTML, ">")
				ElseIf strType = "Attribute" Then
					Dim intStart, _
						intEnd

					intStart = InStr(objChildElements(j).outerHTML, "<") + Len(objChildElements(j).tagName) + 1
					intEnd = InStr(objChildElements(j).outerHTML, ">")

					intInputLimit = intEnd - intStart

					If Not intStart = intEnd Then
						strInput = Mid(objChildElements(j).outerHTML, intStart, intInputLimit)
					Else
						strInput = ""
					End If
				End If

				If objRegex.Test(strInput) Then
					Set objMatches = objRegex.Execute(strInput)
					
					For k = 0 To objMatches.Count - 1
						If objMatches(k).FirstIndex < intInputLimit Then
							objResultSet.Append objChildElements(j)
							Exit For
						End If
					Next
				End If
			Next
		Next

		Set MatchAll = HTMLParser(objResultSet)
	End Function


	' Helper Methods


	' *** The SanitizeHTML() method is similar to lxml.html.clean:

	' The module lxml.html.clean provides a Cleaner class for cleaning up HTML pages.
	' It supports removing embedded or script content, special tags, CSS style annotations
	' and much more. Say, you have an evil web page from an untrusted source that contains
	' lots of content that upsets browsers and tries to run evil code on the client side:

	' To remove the all suspicious content from this unparsed document, use the clean_html
	' function:

	' The Cleaner class supports several keyword arguments to control exactly which content
	' is removed:


	' cleaner = Cleaner(page_structure=False, links=False)


	' cleaner = Cleaner(style=True, links=True, add_nofollow=True,
	' ... page_structure=False, safe_attrs_only=False)


	Private Function SanitizeHTML(strHTML)
		Dim meta_replace, _
			cdata_script_replace, _
			script_replace, _
			event_listeners_select, _
			event_listeners_replace, _
			linked_elements_select, _
			linked_elements_replace, _
			embedded_elements_select, _
			embedded_elements_replace, _
			object_data_select, _
			object_data_replace, _
			applet_codebase_select, _
			applet_codebase_replace, _
			param_name_select, _
			param_name_replace, _
			style_select, _
			style_replace, _
			style_url_select, _
			style_url_replace, _
			style_url_attr_replace, _
			css_expression_inline_select, _
			css_expression_inline_replace, _
			css_expression_select, _
			css_expression_replace, _
			css_expression_attr_replace, _
			html_replace, _
			prefix, _
			match

		prefix = "<!--""'-->"

		Set meta_replace = New RegExp

		With meta_replace
			.Pattern = "<meta(?:[^<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*[^-a-z0-9:._]http-equiv\s*=" & _
					"\s*(?:\""(?:r|&#0*114(?:;|(?!\d))|&#x0*72(?:;|(?!\d))|&#0*82(?:;|(?!" & _
					"\d))|&#x0*52(?:;|(?!\d)))(?:e|&#0*101(?:;|(?!\d))|&#x0*65(?:;|(?!\d" & _
					"))|&#0*69(?:;|(?!\d))|&#x0*45(?:;|(?!\d)))(?:f|&#0*102(?:;|(?!\d))|" & _
					"&#x0*66(?:;|(?!\d))|&#0*70(?:;|(?!\d))|&#x0*46(?:;|(?!\d)))(?:r|&#0" & _
					"*114(?:;|(?!\d))|&#x0*72(?:;|(?!\d))|&#0*82(?:;|(?!\d))|&#x0*52(?:;" & _
					"|(?!\d)))(?:e|&#0*101(?:;|(?!\d))|&#x0*65(?:;|(?!\d))|&#0*69(?:;|(?" & _
					"!\d))|&#x0*45(?:;|(?!\d)))(?:s|&#0*115(?:;|(?!\d))|&#x0*73(?:;|(?!\" & _
					"d))|&#0*83(?:;|(?!\d))|&#x0*53(?:;|(?!\d)))(?:h|&#0*104(?:;|(?!\d))" & _
					"|&#x0*68(?:;|(?!\d))|&#0*72(?:;|(?!\d))|&#x0*48(?:;|(?!\d)))\""(?:[^" & _
					"<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*(?:>|(?=<))|'(?:r|&#0*114(?:;|(?!" & _
					"\d))|&#x0*72(?:;|(?!\d))|&#0*82(?:;|(?!\d))|&#x0*52(?:;|(?!\d)))(?:" & _
					"e|&#0*101(?:;|(?!\d))|&#x0*65(?:;|(?!\d))|&#0*69(?:;|(?!\d))|&#x0*4" & _
					"5(?:;|(?!\d)))(?:f|&#0*102(?:;|(?!\d))|&#x0*66(?:;|(?!\d))|&#0*70(?" & _
					":;|(?!\d))|&#x0*46(?:;|(?!\d)))(?:r|&#0*114(?:;|(?!\d))|&#x0*72(?:;|" & _
					"(?!\d))|&#0*82(?:;|(?!\d))|&#x0*52(?:;|(?!\d)))(?:e|&#0*101(?:;|(?!" & _
					"\d))|&#x0*65(?:;|(?!\d))|&#0*69(?:;|(?!\d))|&#x0*45(?:;|(?!\d)))(?:s" & _
					"|&#0*115(?:;|(?!\d))|&#x0*73(?:;|(?!\d))|&#0*83(?:;|(?!\d))|&#x0*53(" & _
					"?:;|(?!\d)))(?:h|&#0*104(?:;|(?!\d))|&#x0*68(?:;|(?!\d))|&#0*72(?:;|" & _
					"(?!\d))|&#x0*48(?:;|(?!\d)))'(?:[^<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*" & _
					"(?:>|(?=<))|(?:r|&#0*114(?:;|(?!\d))|&#x0*72(?:;|(?!\d))|&#0*82(?:;|" & _
					"(?!\d))|&#x0*52(?:;|(?!\d)))(?:e|&#0*101(?:;|(?!\d))|&#x0*65(?:;|(?!" & _
					"\d))|&#0*69(?:;|(?!\d))|&#x0*45(?:;|(?!\d)))(?:f|&#0*102(?:;|(?!\d))" & _
					"|&#x0*66(?:;|(?!\d))|&#0*70(?:;|(?!\d))|&#x0*46(?:;|(?!\d)))(?:r|&#0" & _
					"*114(?:;|(?!\d))|&#x0*72(?:;|(?!\d))|&#0*82(?:;|(?!\d))|&#x0*52(?:;|" & _
					"(?!\d)))(?:e|&#0*101(?:;|(?!\d))|&#x0*65(?:;|(?!\d))|&#0*69(?:;|(?!\" & _
					"d))|&#x0*45(?:;|(?!\d)))(?:s|&#0*115(?:;|(?!\d))|&#x0*73(?:;|(?!\d))" & _
					"|&#0*83(?:;|(?!\d))|&#x0*53(?:;|(?!\d)))(?:h|&#0*104(?:;|(?!\d))|&#x" & _
					"0*68(?:;|(?!\d))|&#0*72(?:;|(?!\d))|&#x0*48(?:;|(?!\d)))(?:(?:\s|&nb" & _
					"sp;?|&#0*32(?:;|(?!\d))|&#x0*20(?:;|(?!\d)))(?:[^<>""']*(?:""[^""]*""|'[" & _
					"^']*'))*?[^<>]*(?:>|(?=<))|(?:>|(?=<))))"

			.Global = True
			.IgnoreCase = True
		End With

		strHTML = meta_replace.Replace(strHTML, "<!-- meta http-equiv=refresh stripped-->")

		Set cdata_script_replace = New RegExp

		With cdata_script_replace
			.Pattern = "<script(?:[^<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*>\s*//\s*<\[CDATA\[[\S\s]*?]]>\s" & _
				"*</script[^>]*>"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = cdata_script_replace.Replace(strHTML, "<!--CDATA script-->")

		Set script_replace = New RegExp

		With script_replace
			.Pattern = "<script[\S\s]+?<\/script\s*>"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = script_replace.Replace(strHTML, "<!--Non-CDATA script-->")

		Set event_listeners_select = New RegExp

		With event_listeners_select
			.Pattern = "(<[a-z](?:[^<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*[^-a-z0-9:._]on[-a-z0-9:_.]+=(?:[^" & _
					"<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*(?:>|(?=<)))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = event_listeners_select.Replace(strHTML, prefix & "$1")

		Set event_listeners_replace = New RegExp

		With event_listeners_replace
			.Pattern = "([^-a-z0-9:._])(on[-a-z0-9:_.]+\s*=(?:\s*""[^""]*""|\s*'[^']*'|[^\s]+))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = event_listeners_replace.Replace(strHTML, "$1" & "data-" & "$2") 

		Set linked_elements_select = New RegExp

		With linked_elements_select
			.Pattern = "(<[a-z](?:[^<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*[^-a-z0-9:._]href\s*=(?:[^<>""']*(?:""[^""]" & _
					"*""|'[^']*'))*?[^<>]*(?:>|(?=<)))"
			.Global = True
			.IgnoreCase = True
		End With	

		strHTML = linked_elements_select.Replace(strHTML, prefix & "$1")

		Set linked_elements_replace = New RegExp

		With linked_elements_replace
			.Pattern = "([^-a-z0-9:._])(href\s*=(?:\s*""[^""]*""|\s*'[^']*'|[^\s]+))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = linked_elements_replace.Replace(strHTML, "$1" & "data-" & "$2")

		Set embedded_elements_select = New RegExp

		With embedded_elements_select
			.Pattern = "(<[a-z](?:[^<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*[^-a-z0-9:._]src\s*=(?:[^<>""']*(?:""[^""]*""|'" & _
				"[^']*'))*?[^<>]*(?:>|(?=<)))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = embedded_elements_select.Replace(strHTML, prefix & "$1")

		Set embedded_elements_replace = New RegExp

		With embedded_elements_replace
			.Pattern = "([^-a-z0-9:._])(src\s*=(?:\s*""[^""]*""|\s*'[^']*'|[^\s]+))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = embedded_elements_replace.Replace(strHTML, "$1" & "data-" & "$2")

		Set object_data_select = New RegExp

		With object_data_select
			.Pattern = "(<object(?:[^<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*[^-a-z0-9:._]data\s*=" & _
					"(?:[^<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*(?:>|(?=<)))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = object_data_select.Replace(strHTML, prefix & "$1")

		Set object_data_replace = New RegExp

		With object_data_replace
			.Pattern = "([^-a-z0-9:._])(data\s*=(?:\s*""[^""]*""|\s*'[^']*'|[^\s]+))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = object_data_replace.Replace(strHTML, "$1" & "data-" & "$2")

		Set applet_codebase_select = New RegExp

		With applet_codebase_select
			.Pattern = "(<applet(?:[^<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*[^-a-z0-9:._]codebase\s*" & _
					"=(?:[^<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*(?:>|(?=<)))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = applet_codebase_select.Replace(strHTML, prefix & "$1")

		Set applet_codebase_replace = New RegExp

		With applet_codebase_replace
			.Pattern = "([^-a-z0-9:._])(codebase\s*=(?:\s*""[^""]*""|\s*'[^']*'|[^\s]+))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = applet_codebase_replace.Replace(strHTML, "$1" & "data-" & "$2")

		Set param_name_select = New RegExp

		With param_name_select
			.Pattern = "(<param(?:[^<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*[^-a-z0-9:._]name\s*=\s*(?:\""(?" & _
					":m|&#0*109(?:;|(?!\d))|&#x0*6D(?:;|(?!\d))|&#0*77(?:;|(?!\d))|&#x0*4D(?:" & _
					";|(?!\d)))(?:o|&#0*111(?:;|(?!\d))|&#x0*6F(?:;|(?!\d))|&#0*79(?:;|(?!\d)" & _
					")|&#x0*4F(?:;|(?!\d)))(?:v|&#0*118(?:;|(?!\d))|&#x0*76(?:;|(?!\d))|&#0*8" & _
					"6(?:;|(?!\d))|&#x0*56(?:;|(?!\d)))(?:i|&#0*105(?:;|(?!\d))|&#x0*69(?:;|" & _
					"(?!\d))|&#0*73(?:;|(?!\d))|&#x0*49(?:;|(?!\d)))(?:e|&#0*101(?:;|(?!\d))|" & _
					"&#x0*65(?:;|(?!\d))|&#0*69(?:;|(?!\d))|&#x0*45(?:;|(?!\d)))\""(?:[^<>""']*" & _
					"(?:""[^""]*""|'[^']*'))*?[^<>]*(?:>|(?=<))|'(?:m|&#0*109(?:;|(?!\d))|&#x0*6" & _
					"D(?:;|(?!\d))|&#0*77(?:;|(?!\d))|&#x0*4D(?:;|(?!\d)))(?:o|&#0*111(?:;|(?" & _
					"!\d))|&#x0*6F(?:;|(?!\d))|&#0*79(?:;|(?!\d))|&#x0*4F(?:;|(?!\d)))(?:v|&#" & _
					"0*118(?:;|(?!\d))|&#x0*76(?:;|(?!\d))|&#0*86(?:;|(?!\d))|&#x0*56(?:;|(?!" & _
					"\d)))(?:i|&#0*105(?:;|(?!\d))|&#x0*69(?:;|(?!\d))|&#0*73(?:;|(?!\d))|&#x" & _
					"0*49(?:;|(?!\d)))(?:e|&#0*101(?:;|(?!\d))|&#x0*65(?:;|(?!\d))|&#0*69(?:;" & _
					"|(?!\d))|&#x0*45(?:;|(?!\d)))'(?:[^<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*(?:" & _
					">|(?=<))|(?:m|&#0*109(?:;|(?!\d))|&#x0*6D(?:;|(?!\d))|&#0*77(?:;|(?!\d))" & _
					"|&#x0*4D(?:;|(?!\d)))(?:o|&#0*111(?:;|(?!\d))|&#x0*6F(?:;|(?!\d))|&#0*79" & _
					"(?:;|(?!\d))|&#x0*4F(?:;|(?!\d)))(?:v|&#0*118(?:;|(?!\d))|&#x0*76(?:;|(?" & _
					"!\d))|&#0*86(?:;|(?!\d))|&#x0*56(?:;|(?!\d)))(?:i|&#0*105(?:;|(?!\d))|&#" & _
					"x0*69(?:;|(?!\d))|&#0*73(?:;|(?!\d))|&#x0*49(?:;|(?!\d)))(?:e|&#0*101(?:" & _
					";|(?!\d))|&#x0*65(?:;|(?!\d))|&#0*69(?:;|(?!\d))|&#x0*45(?:;|(?!\d)))(?:" & _
					"(?:\s|&nbsp;?|&#0*32(?:;|(?!\d))|&#x0*20(?:;|(?!\d)))(?:[^<>""']*(?:""[^""]" & _
					"*""|'[^']*'))*?[^<>]*(?:>|(?=<))|(?:>|(?=<)))))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = param_name_select.Replace(strHTML, prefix & "$1")

		Set param_name_replace = New RegExp

		With param_name_replace
			.Pattern = "([^-a-z0-9:._])(value\s*=(?:\s*""[^""]*""|\s*'[^']*'|[^\s]+))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = param_name_replace.Replace(strHTML, "$1" & "data-" & "$2")

		Set style_select = New RegExp

		With style_select
			.Pattern = "(<style[^>]*>(?:[^""']*(?:""[^""]*""|'[^']*'))*?[^'""]*(?:<\/style|$))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = style_select.Replace(strHTML, prefix & "$1")

		Set style_replace = New RegExp

		With style_replace
			.Pattern = "([^-a-z0-9:._])(url\s*\(\s*(?:\s*""[^""]*""|\s*'[^']*'|[^\s]+?)\s*\))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = style_replace.Replace(strHTML, "$1" & "data-" & "$2")

		Set style_url_select = New RegExp

		With style_url_select
			.Pattern = "(<[a-z](?:[^<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*[^-a-z0-9:._]style\s*=(?:[^<>""']*" & _
					"(?:""[^""]*""|'[^']*'))*?[^<>]*(?:>|(?=<)))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = style_url_select.Replace(strHTML, prefix & "$1")

		Set style_url_replace = New RegExp

		With style_url_replace
			.Pattern = "([^-a-z0-9:._]style\s*=)((?:\s*\""[^\""]*\""|\s*'[^']*'|[^\s>]+))"
			.Global = True
			.IgnoreCase = True
		End With

		Set style_url_attr_replace = New RegExp

		With style_url_attr_replace
			For Each match in style_url_replace.Execute(strHTML)	
				If Mid(match.SubMatches(1), 1, 1) = """" Then
					.Pattern = "("")((?:u|&#0*117(?:;|(?!\d))|&#x0*75(?:;|(?!\d))|&#0*85(?:;|(?!\d))|&" & _
							"#x0*55(?:;|(?!\d)))(?:r|&#0*114(?:;|(?!\d))|&#x0*72(?:;|(?!\d))|" & _
							"&#0*82(?:;|(?!\d))|&#x0*52(?:;|(?!\d)))(?:l|&#0*108(?:;|(?!\d))|" & _
							"&#x0*6C(?:;|(?!\d))|&#0*76(?:;|(?!\d))|&#x0*4C(?:;|(?!\d)))(?:\(" & _
							"|&#0*40(?:;|(?!\d))|&#x0*28(?:;|(?!\d)))[^""]+"")"
				ElseIf Mid(match.SubMatches(1), 1, 1) = "'" Then
					.Pattern = "(')((?:u|&#0*117(?:;|(?!\d))|&#x0*75(?:;|(?!\d))|&#0*85(?:;|(?!\d))|&" & _
							"#x0*55(?:;|(?!\d)))(?:r|&#0*114(?:;|(?!\d))|&#x0*72(?:;|(?!\d))|" & _
							"&#0*82(?:;|(?!\d))|&#x0*52(?:;|(?!\d)))(?:l|&#0*108(?:;|(?!\d))|" & _
							"&#x0*6C(?:;|(?!\d))|&#0*76(?:;|(?!\d))|&#x0*4C(?:;|(?!\d)))(?:\(" & _
							"|&#0*40(?:;|(?!\d))|&#x0*28(?:;|(?!\d)))[^']+')"
				Else
					.Pattern = "()((?:u|&#0*117(?:;|(?!\d))|&#x0*75(?:;|(?!\d))|&#0*85(?:;|(?!\d))|&#" & _
							"x0*55(?:;|(?!\d)))(?:r|&#0*114(?:;|(?!\d))|&#x0*72(?:;|(?!\d))|&" & _
							"#0*82(?:;|(?!\d))|&#x0*52(?:;|(?!\d)))(?:l|&#0*108(?:;|(?!\d))|&" & _
							"#x0*6C(?:;|(?!\d))|&#0*76(?:;|(?!\d))|&#x0*4C(?:;|(?!\d)))(?:\(|" & _
							"&#0*40(?:;|(?!\d))|&#x0*28(?:;|(?!\d)))(?:""[^""]+""|\'[^\']+\'|(?:" & _
							"(?!(?:\)|&#0*41(?:;|(?!\d))|&#x0*29(?:;|(?!\d)))).)+)(?:\)|&#0*4" & _
							"1(?:;|(?!\d))|&#x0*29(?:;|(?!\d))))"
				End If

				strHTML = .Replace(strHTML, "$1" & "data-" & "$2")
			Next
		End With

		Set css_expression_inline_select = New RegExp

		With css_expression_inline_select
			.Pattern = "(<style[^>]*>(?:[^""']*(?:""[^""]*""|'[^']*'))*?[^'""]*(?:<\/style|$))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = css_expression_inline_select.Replace(strHTML, prefix & "$1")

		Set css_expression_inline_replace = New RegExp

		With css_expression_inline_replace
			.Pattern = "([^-a-z0-9:._])(expression\s*\(\s*(?:\s*""[^""]*""|\s*'[^']*'|[^\s]+?)\s*\))"
			.Global = True
			.IgnoreCase = True
		End With

		Set css_expression_select = New RegExp

		With css_expression_select
			.Pattern = "(<[a-z](?:[^<>""']*(?:""[^""]*""|'[^']*'))*?[^<>]*[^-a-z0-9:._]style\s*=(?:[^<>""']*(?:""[^""]" & _
					"*""|'[^']*'))*?[^<>]*(?:>|(?=<)))"
			.Global = True
			.IgnoreCase = True
		End With

		Set css_expression_replace = New RegExp

		With css_expression_replace
			.Pattern = "([^-a-z0-9:._]style\s*=)((?:\s*\""[^\""]*\""|\s*'[^']*'|[^\s>]+))"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = style_replace.Replace(strHTML, "$1" & "data-" & "$2")

		Set css_expression_attr_replace = New RegExp
	
		With css_expression_attr_replace
			For Each match in css_expression_replace.Execute(strHTML)	
				If Mid(match.SubMatches(1), 1, 1) = """" Then
					.Pattern = "("")((?:e|&#0*101(?:;|(?!\d))|&#x0*65(?:;|(?!\d))|&#0*69(?:;|(?!\d))|&" & _
							"#x0*45(?:;|(?!\d)))(?:x|&#0*120(?:;|(?!\d))|&#x0*78(?:;|(?!\d))|" & _
							"&#0*88(?:;|(?!\d))|&#x0*58(?:;|(?!\d)))(?:p|&#0*112(?:;|(?!\d))|" & _
							"&#x0*70(?:;|(?!\d))|&#0*80(?:;|(?!\d))|&#x0*50(?:;|(?!\d)))(?:r|" & _
							"&#0*114(?:;|(?!\d))|&#x0*72(?:;|(?!\d))|&#0*82(?:;|(?!\d))|&#x0*" & _
							"52(?:;|(?!\d)))(?:e|&#0*101(?:;|(?!\d))|&#x0*65(?:;|(?!\d))|&#0*" & _
							"69(?:;|(?!\d))|&#x0*45(?:;|(?!\d)))(?:s|&#0*115(?:;|(?!\d))|&#x0" & _
							"*73(?:;|(?!\d))|&#0*83(?:;|(?!\d))|&#x0*53(?:;|(?!\d)))(?:s|&#0*" & _
							"115(?:;|(?!\d))|&#x0*73(?:;|(?!\d))|&#0*83(?:;|(?!\d))|&#x0*53(?" & _
							":;|(?!\d)))(?:i|&#0*105(?:;|(?!\d))|&#x0*69(?:;|(?!\d))|&#0*73(?" & _
							":;|(?!\d))|&#x0*49(?:;|(?!\d)))(?:o|&#0*111(?:;|(?!\d))|&#x0*6F(" & _
							"?:;|(?!\d))|&#0*79(?:;|(?!\d))|&#x0*4F(?:;|(?!\d)))(?:n|&#0*110(" & _
							"?:;|(?!\d))|&#x0*6E(?:;|(?!\d))|&#0*78(?:;|(?!\d))|&#x0*4E(?:;|(" & _
							"?!\d)))(?:\(|&#0*40(?:;|(?!\d))|&#x0*28(?:;|(?!\d)))[^""]+"")"
				ElseIf Mid(match.SubMatches(1), 1, 1) = "'" Then
					.Pattern = "(')((?:e|&#0*101(?:;|(?!\d))|&#x0*65(?:;|(?!\d))|&#0*69(?:;|(?!\d))|&" & _
							"#x0*45(?:;|(?!\d)))(?:x|&#0*120(?:;|(?!\d))|&#x0*78(?:;|(?!\d))|" & _
							"&#0*88(?:;|(?!\d))|&#x0*58(?:;|(?!\d)))(?:p|&#0*112(?:;|(?!\d))|" & _
							"&#x0*70(?:;|(?!\d))|&#0*80(?:;|(?!\d))|&#x0*50(?:;|(?!\d)))(?:r|" & _
							"&#0*114(?:;|(?!\d))|&#x0*72(?:;|(?!\d))|&#0*82(?:;|(?!\d))|&#x0*" & _
							"52(?:;|(?!\d)))(?:e|&#0*101(?:;|(?!\d))|&#x0*65(?:;|(?!\d))|&#0*" & _
							"69(?:;|(?!\d))|&#x0*45(?:;|(?!\d)))(?:s|&#0*115(?:;|(?!\d))|&#x0" & _
							"*73(?:;|(?!\d))|&#0*83(?:;|(?!\d))|&#x0*53(?:;|(?!\d)))(?:s|&#0*" & _
							"115(?:;|(?!\d))|&#x0*73(?:;|(?!\d))|&#0*83(?:;|(?!\d))|&#x0*53(?" & _
							":;|(?!\d)))(?:i|&#0*105(?:;|(?!\d))|&#x0*69(?:;|(?!\d))|&#0*73(?" & _
							":;|(?!\d))|&#x0*49(?:;|(?!\d)))(?:o|&#0*111(?:;|(?!\d))|&#x0*6F(?" & _
							":;|(?!\d))|&#0*79(?:;|(?!\d))|&#x0*4F(?:;|(?!\d)))(?:n|&#0*110(?:" & _
							";|(?!\d))|&#x0*6E(?:;|(?!\d))|&#0*78(?:;|(?!\d))|&#x0*4E(?:;|(?!" & _
							"\d)))(?:\(|&#0*40(?:;|(?!\d))|&#x0*28(?:;|(?!\d)))[^']+')"
				Else
					.Pattern = "()((?:e|&#0*101(?:;|(?!\d))|&#x0*65(?:;|(?!\d))|&#0*69(?:;|(?!\d))|&#x" & _
							"0*45(?:;|(?!\d)))(?:x|&#0*120(?:;|(?!\d))|&#x0*78(?:;|(?!\d))|&#0" & _
							"*88(?:;|(?!\d))|&#x0*58(?:;|(?!\d)))(?:p|&#0*112(?:;|(?!\d))|&#x0" & _
							"*70(?:;|(?!\d))|&#0*80(?:;|(?!\d))|&#x0*50(?:;|(?!\d)))(?:r|&#0*1" & _
							"14(?:;|(?!\d))|&#x0*72(?:;|(?!\d))|&#0*82(?:;|(?!\d))|&#x0*52(?:;" & _
							"|(?!\d)))(?:e|&#0*101(?:;|(?!\d))|&#x0*65(?:;|(?!\d))|&#0*69(?:;|" & _
							"(?!\d))|&#x0*45(?:;|(?!\d)))(?:s|&#0*115(?:;|(?!\d))|&#x0*73(?:;|" & _
							"(?!\d))|&#0*83(?:;|(?!\d))|&#x0*53(?:;|(?!\d)))(?:s|&#0*115(?:;|(" & _
							"?!\d))|&#x0*73(?:;|(?!\d))|&#0*83(?:;|(?!\d))|&#x0*53(?:;|(?!\d))" & _
							")(?:i|&#0*105(?:;|(?!\d))|&#x0*69(?:;|(?!\d))|&#0*73(?:;|(?!\d))|" & _
							"&#x0*49(?:;|(?!\d)))(?:o|&#0*111(?:;|(?!\d))|&#x0*6F(?:;|(?!\d))|" & _
							"&#0*79(?:;|(?!\d))|&#x0*4F(?:;|(?!\d)))(?:n|&#0*110(?:;|(?!\d))|&" & _
							"#x0*6E(?:;|(?!\d))|&#0*78(?:;|(?!\d))|&#x0*4E(?:;|(?!\d)))(?:\(|&" & _
							"#0*40(?:;|(?!\d))|&#x0*28(?:;|(?!\d)))(?:""[^""]+""|\'[^\']+\'|(?:(?" & _
							"!(?:\)|&#0*41(?:;|(?!\d))|&#x0*29(?:;|(?!\d)))).)+)(?:\)|&#0*41(?" & _
							":;|(?!\d))|&#x0*29(?:;|(?!\d))))"
				End If

				strHTML = .Replace(strHTML, "$1" & "data-" & "$2")
			Next
		End With

		Set html_replace = New RegExp

		With html_replace
			.Pattern = "(?:<!--""'-->)+"
			.Global = True
			.IgnoreCase = True
		End With

		strHTML = html_replace.Replace(strHTML, "<!--""'-->")

		SanitizeHTML = strHTML
	End Function

	Private Function HTMLParser(varHTML)
		Dim objParser, _
			objArr, _
			i

		Set objParser = New base_HTML_Parser
		Set objArr = New base_Data_Array

		If TypeName(varHTML) = "base_Data_Array" Then
			Set objArr = varHTML
		ElseIf TypeName(varHTML) = "DispHTMLElementCollection" Then
			For i = 0 To varHTML.length - 1
				objArr.Append varHTML(i)
			Next
		ElseIf varHTML.nodeType = 1 Or varHTML.nodeType = 3 Or varHTML.nodeType = 8 Then
			objArr.Append varHTML
		End If

		objParser.SelectionFrom pHtmlParser, objArr

		Set HTMLParser = objParser
	End Function

	Public Sub Load(strFile)
		Dim objFSO
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		FromString objFSO.OpenTextFile(strFile, 1).ReadAll()
		Set FSO = Nothing
	End Sub

	Public Sub FromString(strHTML)
		If pSanitize Then strHTML = SanitizeHTML(strHTML)
		pHtmlParser.WriteLn strHTML
		ClearSelection()
		pSelection.Append pHtmlParser.getElementsByTagName("html")(0)
	End Sub

	Public Function ToString()
		ToString = pHtmlParser.documentElement.outerHTML
	End Function

	Public Sub FromDocument(objHtmlDoc)
		If objHtmlDoc.nodeType = 9 Then
			Set pHtmlParser = objHtmlDoc.cloneNode(True)
			ClearSelection()
			pSelection.Append pHtmlParser.getElementsByTagName("html")(0)
		End If
	End Sub

	Public Function ToDocument()
		Set ToDocument = pHtmlParser
	End Function

	Public Function SelectionFrom(objHtmlParser, objArr)
		If TypeName(objHtmlParser) = "HTMLDocument" Then
			If objHtmlParser.nodeType = 9 Then Set pHtmlParser = objHtmlParser
		End If
		If TypeName(objArr) = "base_Data_Array" Then Set pSelection = objArr
	End Function

	Public Sub Clear()
		Class_Initialize()
	End Sub

	Public Sub ClearSelection()
		pSelection.Clear()
	End Sub

	Private Sub Class_Terminate()
		Set pHtmlParser = Nothing
		Set pSelection = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_HTML_Parser.vbs" Then
	Dim parser
	Set parser = New base_HTML_Parser

	parser.Sanitize = True
	parser.FromString "<html><head><title>This is the page's title.</title>" & _
				"<meta name=""author"" content=""name"">" & _
				"<meta name=""keywords"" content=""php documentation"">" & _
				"<meta name=""DESCRIPTION"" content=""a php manual"">" & _
				"<meta name=""geo.position"" content=""49.33;-86.59"">" & _
				"<!-- This is the header area -->" & _
				"<link rel=""stylesheet"" href=""styles.css"">" & _
				"<style>" & _
				"h1 {color:red;}" & _
				"p {color:blue;}" & _
				"</style>" & _
				"<style></style>" & _
				"</head>" & _
				"<body><div id=""testID""><h1>Page Title</h1><p class=""testClass"">This is an example" & _
				" paragraph.</p><div><p>You should check this out:</p></div></div><p id=""test"" class=""testClass"">Another test paragraph.</p>" & _
				"<form>First name:<br><input type=""text"" name=""firstname""><br>" & _
  				"Last name:<br><input type=""text"" name=""lastname""></form>" & _
				"<a href=""http://www.w3schools.com"">Visit W3Schools.com!</a>" & _
				"<a href=""http://www.startech.com"">Visit StarTech.com!</a>" & _
				"<img src=""smiley.gif"" alt=""Smiley face"" height=""42"" width=""42"">" & _
  				"<frameset>" & _
				"<frame src=""https://developer.mozilla.org/en/HTML/Element/iframe"" />" & _
  				"<frame src=""https://developer.mozilla.org/en/HTML/Element/frame"" />" & _
				"</frameset>" & _
				"<script>" & _
				"document.getElementById(""testID"").innerHTML = ""Hello JavaScript!"";" & _
				"</script><!-- This is a comment -->" & _
				"<span>Testing, testing!</span>" & _
				"<form action=""demo_form.asp"">" & _
				"<label for=""male"">Male</label>" & _
				"<input type=""radio"" name=""gender"" id=""male"" value=""male""><br>" & _
				"<label for=""female"">Female</label>" & _
				"<input type=""radio"" name=""gender"" id=""female"" value=""female""><br>" & _
				"<label for=""other"">Other</label>" & _
				"<input type=""radio"" name=""gender"" id=""other"" value=""other""><br><br>" & _
				"<input type=""submit"" value=""Submit"">" & _
				"<button type=""button"">Click Me!</button>" & _
				"</form>" & _
				"<table>" & _
				"<tr>" & _
				"<th>Month</th>" & _
				"<th>Savings</th>" & _
				"</tr>" & _
				"<tr>" & _
				"<td>January</td>" & _
				"<td>$100</td>" & _
				"</tr>" & _
				"</table>" & _
				"<ul>" & _
				"<li>Coffee</li>" & _
				"<li>Tea</li>" & _
				"<li>Milk</li>" & _
				"</ul>" & _
				"<ol>" & _
				"<li>Coffee</li>" & _
				"<li>Tea</li>" & _
				"<li>Milk</li>" & _
				"</ol>" & _			
				"<p class=""testClass"">A third test paragraph.</p>" & _
				"<textarea name=""animal"">A paragraph about animals!</textarea></body></html>"

	' WScript.Echo parser.ElementByID("testID").FirstChild.Sibling(1).Parent.Child(1).HTML
	' WScript.Echo parser.Body.Selector("#testID").H1.HTML

	WScript.Echo parser.Body.Match("HTML", "<div", True, True).MatchAll("Text", "You", True, True).HTML
End If
