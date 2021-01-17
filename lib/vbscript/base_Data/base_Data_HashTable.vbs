Option Explicit

Class base_Data_HashTable
	Private pHashTable

	Private Sub Class_Initialize()
		Set pHashTable = CreateObject("System.Collections.Hashtable")
	End Sub


	' Properties


	Public Property Get Count()
		Count = pHashTable.Count
	End Property

	Public Property Get IsFixedSize()
		IsFixedSize = pHashTable.IsFixedSize
	End Property

	Public Property Get IsReadOnly()
		IsReadOnly = pHashTable.IsReadOnly
	End Property

	Public Property Get IsSynchronized()
		IsSynchronized = pHashTable.IsSynchronized
	End Property

	Public Default Property Get Item(intIndex)
		If IsObject(pHashTable(intIndex)) Then
			Set Item = pHashTable(intIndex)
		Else
			Item = pHashTable(intIndex)
		End If
	End Property

	Public Property Let Item(intIndex, varInput)
		pHashTable(intIndex) = varInput
	End Property

	Public Property Set Item(intIndex, varInput)
		Set pHashTable(intIndex) = varInput
	End Property

	Public Property Get Keys()
		Set Keys = pHashTable.Keys
	End Property

	Public Property Get SyncRoot()
		SyncRoot = pHashTable.SyncRoot
	End Property

	Public Property Get Values()
		Set Values = pHashTable.Values
	End Property


	' Methods


	Public Sub Add(strKey, varItem)
		pHashTable.Add strKey, varItem
	End Sub

	Public Sub Clear()
		pHashTable.Clear()
	End Sub

	Public Function Clone()
		Set Clone = pHashTable.Clone()
	End Function

	Public Function Contains(strKey)
		Contains = pHashTable.Contains(strKey)
	End Function

	Public Function ContainsKey(strKey)
		ContainsKey = pHashTable.ContainsKey(strKey)
	End Function

	Public Function ContainsValue(varItem)
		ContainsValue = pHashTable.ContainsValue(varItem)
	End Function

	Public Function Equals(objItem)
		Equals = pHashTable.Equals(objItem)
	End Function

	Public Function GetEnumerator()
		Set GetEnumerator = pHashTable.GetEnumerator()
	End Function

	Public Function GetHashCode()
		GetHashCode = pHashTable.GetHashCode()
	End Function

	Public Function GetType()
		GetType = pHashTable.GetType()
	End Function

	Public Function Remove(strKey)
		pHashTable.Remove strKey
	End Function

	Public Function ToString()
		ToString = pHashTable.ToString()
	End Function

	Private Sub Class_Terminate()
		Set pHashTable = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Data_HashTable.vbs" Then
	Dim hash
	Set hash = New base_Data_HashTable

	hash.Add "FirstName", "Sam"
	hash.Add "LastName", "Smith"
	hash.Add "Title", "Supervisor"
	hash.Add "EmployeeCode", 1457345

	WScript.Echo hash("Title")
End If
