Option Explicit

Class base_Data_Dictionary
	Private pDictionary

	Private Sub Class_Initialize()
		Set pDictionary = CreateObject("Scripting.Dictionary")
	End Sub


	' Properties


	Public Property Get CompareMode()
		CompareMode = pDictionary.CompareMode
	End Property

	Public Property Let CompareMode(intCompareMode)
		If pDictionary.Count = 0 Then pDictionary.CompareMode = intCompareMode
	End Property

	Public Property Get Count()
		Count = pDictionary.Count
	End Property
	
	Public Default Property Get Item(strKey)
		If IsObject(pDictionary.Item(strKey)) Then
			Set Item = pDictionary.Item(strKey)
		Else
			Item = pDictionary.Item(strKey)
		End If
	End Property

	Public Property Let Item(strKey, varItem)
		pDictionary.Item(strKey) = varItem
	End Property

	Public Property Set Item(strKey, varItem)
		Set pDictionary.Item(strKey) = varItem
	End Property
	
	Public Property Get Items()
		Items = pDictionary.Items()
	End Property

	Public Property Let Key(strKey, strNewKey)
		pDictionary.Key(strKey) = strNewKey
	End Property

	Public Property Get Keys()
		Keys = pDictionary.Keys()
	End Property


	' Methods


	Public Sub Add(strKey, varItem)
		pDictionary.Add strKey, varItem
	End Sub   

	Public Function Exists(strKey)
		Exists = pDictionary.Exists(strKey)
	End Function
   
	Public Sub Remove(strKey)
		pDictionary.Remove strKey
	End Sub

	Public Sub RemoveAll()
		pDictionary.RemoveAll
	End Sub

	Private Sub Class_Terminate()
		Set pDictionary = Nothing
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
