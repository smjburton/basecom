Option Explicit

Class v_Data_List
	Private pList

	Private Sub Class_Initialize()
		Set pList = CreateObject("System.Collections.SortedList")
	End Sub


	' Properties


	Public Property Get Capacity()
		Capacity = pList.Capacity
	End Property

	Public Property Let Capacity(intCapacity)
		pList.Capacity = intCapacity
	End Property

	Public Property Get Count()
		Count = pList.Count
	End Property

	Public Property Get IsFixedSize()
		IsFixedSize = pList.IsFixedSize
	End Property

	Public Property Get IsReadOnly()
		IsReadOnly = pList.IsReadOnly
	End Property

	Public Property Get IsSynchronized()
		IsSynchronized = pList.IsSynchronized
	End Property

	Public Default Property Get Item(strKey)
		If IsObject(pList.Item(strKey)) Then
			Set Item = pList.Item(strKey)
		Else
			Item = pList.Item(strKey)
		End If
	End Property

	Public Property Let Item(strKey, varItem)
		pList.Item(strKey) = varItem
	End Property

	Public Property Set Item(strKey, varItem)
		Set pList.Item(strKey) = varItem
	End Property

	Public Property Get Keys()
		Set Keys = pList.Keys
	End Property

	Public Property Get SyncRoot()
		SyncRoot = pList.SyncRoot
	End Property

	Public Property Get Values()
		Set Values = pList.Values
	End Property


	' Methods


	Public Sub Add(strKey, varItem)
		pList.Add strKey, varItem
	End Sub

	Public Sub Clear()
		pList.Clear()
	End Sub

	Public Function Clone()
		Set Clone = pList.Clone()
	End Function

	Public Function Contains(varItem)
		Contains = pList.Contains(varItem)
	End Function

	Public Function ContainsKey(strKey)
		ContainsKey = pList.ContainsKey(strKey)
	End Function

	Public Function ContainsValue(varItem)
		ContainsValue = pList.ContainsValue(varItem)
	End Function

	Public Function Equals(objItem)
		Equals = pList.Equals(objItem)
	End Function

	Public Function GetByIndex(intIndex)
		If IsObject(pList.GetByIndex(intIndex)) Then
			Set GetByIndex = pList.GetByIndex(intIndex)
		Else
			GetByIndex = pList.GetByIndex(intIndex)
		End If
	End Function

	Public Function GetEnumerator()
		Set GetEnumerator = pList.GetEnumerator()
	End Function

	Public Function GetHashCode()
		GetHashCode = pList.GetHashCode()
	End Function

	Public Function GetKey(intIndex)
		GetKey = pList.GetKey(intIndex)
	End Function

	Public Function GetKeyList()
		Set GetKeyList = pList.GetKeyList()
	End Function

	Public Function GetType()
		Set GetType = pList.GetType()
	End Function

	Public Function GetValueList()
		Set GetValueList = pList.GetValueList()
	End Function

	Public Function IndexOfKey(strKey)
		IndexOfKey = pList.IndexOfKey(strKey)
	End Function

	Public Function IndexOfValue(varItem)
		IndexOfValue = pList.IndexOfValue(varItem)
	End Function

	Public Sub Remove(varItem)
		pList.Remove(varItem)
	End Sub

	Public Sub RemoveAt(intIndex)
		pList.RemoveAt(intIndex)
	End Sub

	Public Sub SetByIndex(intIndex, varItem)
		pList.SetByIndex intIndex, varItem
	End Sub

	Public Function ToString()
		ToString = pList.ToString()
	End Function

	Public Sub TrimToSize()
		pList.TrimToSize()
	End Sub

	Private Sub Class_Terminate()
		Set pList = Nothing
	End Sub
End Class

If WScript.ScriptName = "v_Data_List.vbs" Then
	Dim list
	Set list = New v_Data_List

	list.Add "Point", 1
	list.Add "Point Cloud", 2
	list.Add "Curve", 4
	list.Add "Surface", 8
	list.Add "Polysurface", 16
	list.Add "Mesh", 32 

	list.TrimToSize()

	WScript.Echo list.Capacity
End If