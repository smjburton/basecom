Option Explicit

Class base_Data_Dictionary
	Private p_Dictionary

	Private Sub Class_Initialize()
		Set p_Dictionary = CreateObject("Scripting.Dictionary")
	End Sub


	' Properties


	Public Property Get CompareMode()
		CompareMode = p_Dictionary.CompareMode
	End Property

	Public Property Let CompareMode(intCompareMode)
		If p_Dictionary.Count = 0 Then p_Dictionary.CompareMode = intCompareMode
	End Property

	Public Property Get Count()
		Count = p_Dictionary.Count
	End Property
	
	Public Default Property Get Item(varKey)
		If IsObject(p_Dictionary.Item(varKey)) Then
			Set Item = p_Dictionary.Item(varKey)
		Else
			Item = p_Dictionary.Item(varKey)
		End If
	End Property

	Public Property Let Item(varKey, varItem)
		p_Dictionary.Item(varKey) = varItem
	End Property

	Public Property Set Item(varKey, varItem)
		Set p_Dictionary.Item(varKey) = varItem
	End Property
	
	Public Property Get Items()
		Items = p_Dictionary.Items()
	End Property

	Public Property Let Key(varKey, varNewKey)
		p_Dictionary.Key(varKey) = varNewKey
	End Property

	Public Property Set Key(varKey, varNewKey)
		p_Dictionary.Key(varKey) = varNewKey
	End Property

	Public Property Get Keys()
		Keys = p_Dictionary.Keys()
	End Property


	' Methods


	Public Sub Add(varKey, varItem)
		p_Dictionary.Add varKey, varItem
	End Sub   

	Public Function Exists(varKey)
		Exists = p_Dictionary.Exists(varKey)
	End Function
   
	Public Sub Remove(varKey)
		p_Dictionary.Remove varKey
	End Sub

	Public Sub RemoveAll()
		p_Dictionary.RemoveAll
	End Sub

	Private Sub Class_Terminate()
		Set p_Dictionary = Nothing
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
