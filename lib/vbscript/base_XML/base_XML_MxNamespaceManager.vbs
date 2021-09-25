Option Explicit

Class base_XML_MxNamespaceManager
	Private p_objMxNamespaceManager

	Private Sub Class_Initialize()
		Set p_objMxNamespaceManager = CreateObject("MSXML2.MXNamespaceManager.6.0")
	End Sub


	' Properties


	Public Property Get AllowOverride()
		AllowOverride = p_objMxNamespaceManager.allowOverride 
	End Property

	Public Property Let AllowOverride( _
		ByVal blnAllowOverride _
		)

		p_objMxNamespaceManager.allowOverride = blnAllowOverride
	End Property


	' Methods


	Public Sub DeclarePrefix( _
		ByVal strPrefix, _
		ByVal strNamespaceUri _
		)

		p_objMxNamespaceManager.declarePrefix strPrefix, strNamespaceUri
	End Sub

	Public Function GetDeclaredPrefixes()
		Set GetDeclaredPrefixes = p_objMxNamespaceManager.getDeclaredPrefixes()
	End Function

	Public Function GetPrefixes( _
		ByVal strNamespaceUri _
		)

		Set GetPrefixes = p_objMxNamespaceManager.getPrefixes(strNamespaceUri)
	End Function

	Public Function GetUri( _
		ByVal strPrefix _
		)

		GetUri = p_objMxNamespaceManager.getURI(strPrefix)
	End Function

	Public Function GetUriFromNode( _
		ByVal strPrefix, _
		ByVal objContextNode _
		)

		GetUriFromNode = p_objMxNamespaceManager.getURIFromNode(strPrefix, objContextNode)
	End Function

	Public Sub PopContext()
		p_objMxNamespaceManager.popContext
	End Sub

	Public Sub PushContext()
		p_objMxNamespaceManager.pushContext
	End Sub

	Public Sub PushNodeContext( _
		ByVal objContextNode _
		) ' Optional params: [fDeep As Boolean = True]

		p_objMxNamespaceManager.pushNodeContext objContextNode
	End Sub

	Public Sub Reset()
		p_objMxNamespaceManager.reset
	End Sub

	Private Sub Class_Terminate()
		Set p_objMxNamespaceManager = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_XML_MxNamespaceManager.vbs" Then

End If