Option Explicit

Class base_Data_Dictionary
	Private p_objDictionary

	Private Sub Class_Initialize()
		Set p_objDictionary = CreateObject("Scripting.Dictionary")
	End Sub


	' Properties


	Public Property Get CompareMode()
		CompareMode = p_objDictionary.CompareMode
	End Property

	Public Property Let CompareMode(intCompareMode)
		If p_objDictionary.Count = 0 Then p_objDictionary.CompareMode = intCompareMode
	End Property

	Public Property Get Count()
		Count = p_objDictionary.Count
	End Property
	
	Public Default Property Get Item(varKey)
		If IsObject(p_objDictionary.Item(varKey)) Then
			Set Item = p_objDictionary.Item(varKey)
		Else
			Item = p_objDictionary.Item(varKey)
		End If
	End Property

	Public Property Let Item(varKey, varItem)
		p_objDictionary.Item(varKey) = varItem
	End Property

	Public Property Set Item(varKey, varItem)
		Set p_objDictionary.Item(varKey) = varItem
	End Property
	
	Public Property Get Items()
		Items = p_objDictionary.Items()
	End Property

	Public Property Let Key(varKey, varNewKey)
		p_objDictionary.Key(varKey) = varNewKey
	End Property

	Public Property Set Key(varKey, varNewKey)
		p_objDictionary.Key(varKey) = varNewKey
	End Property

	Public Property Get Keys()
		Keys = p_objDictionary.Keys()
	End Property


	' Methods


	Public Sub Add(varKey, varItem)
		p_objDictionary.Add varKey, varItem
	End Sub   

	Public Function Exists(varKey)
		Exists = p_objDictionary.Exists(varKey)
	End Function
   
	Public Sub Remove(varKey)
		p_objDictionary.Remove varKey
	End Sub

	Public Sub RemoveAll()
		p_objDictionary.RemoveAll
	End Sub

	Private Sub Class_Terminate()
		Set p_objDictionary = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Data_Dictionary.vbs" Then
	Dim dictionary

	Set dictionary = New base_Data_Dictionary

	dictionary.CompareMode = VBDataBaseCompare

	dictionary.Add "Key 1", "Item 1"
	dictionary.Add "Key 2", "Item 2"
	dictionary.Add "Key 3", "Item 3"
	dictionary.Add "Key 4", "Item 4"

	WScript.Echo dictionary.CompareMode
End If
