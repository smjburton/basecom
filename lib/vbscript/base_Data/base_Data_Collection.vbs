Include "base_Data_Array_Util"

Class base_Data_Collection
	Private p_Collection

	Private Sub Class_Initialize()

	End Sub


	' Properties

	
	Public Property Get Count()
		If IsArrayAllocated(p_Collection) Then Count = UBound(p_Collection) + 1
	End Property

	Public Default Property Get Item(intIndex)
		If IsArrayAllocated(p_Collection) Then
    			If IsObject(p_Collection(intIndex)) Then
        			Set Item = p_Collection(intIndex)
    			Else
        			Item = p_Collection(intIndex)
    			End If
		End If
	End Property


	' Methods


	Public Sub Add(varItem)
		Extend 1

		If IsObject(varItem) Then
			Set p_Collection(UBound(p_Collection)) = varItem
		Else
			p_Collection(UBound(p_Collection)) = varItem
		End If
	End Sub

	Public Sub Clear()
		If IsArrayAllocated(p_Collection) Then Erase p_Collection
	End Sub

	Public Function Contains(varItem)
		Dim blnContains, _
			i

		blnContains = False

		For i = 0 To UBound(p_Collection)
			If IsObject(varItem) And IsObject(p_Collection(i)) Then
				If p_Collection(i) Is varItem Then blnContains = True
			ElseIf Not IsObject(p_Collection(i)) And Not IsObject(varItem) Then
				If p_Collection(i) = varItem Then blnContains = True
			End If
			If blnContains Then Exit For
		Next

		Contains = blnContains
	End Function

	Public Sub Remove(intIndex)
		If IsArrayAllocated(p_Collection) Then
			Dim objCollection, _
				i

			Set objCollection = New v_Data_Collection

			For i = 0 To UBound(p_Collection)
  				Do
    					If i = intIndex Then Exit Do
        				
					objCollection.Add p_Collection(i)
  				Loop While False
			Next

			Clear()

			For i = 0 To objCollection.Count - 1
				Me.Add objCollection(i)
			Next
		End If
	End Sub

	
	' Helper Methods


	Private Sub Extend(intSize)
		If IsArrayAllocated(p_Collection) Then
			ReDim Preserve p_Collection(UBound(p_Collection) + intSize)
		Else
			p_Collection = Array()
			ReDim p_Collection(intSize - 1)
		End If
	End Sub
	
	Private Sub Class_Terminate()

	End Sub
End Class

If WScript.ScriptName = "base_Data_Collection.vbs" Then
	Dim collection, _
		objDict

	Set collection = New base_Data_Collection
	Set objDict = CreateObject("Scripting.Dictionary")

	collection.Add "Car"
	collection.Add True
	collection.Add 342
	collection.Add objDict

	WScript.Echo collection(2)
End If
