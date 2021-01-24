Option Explicit

Include "base_Data.base_Data_Array_Util"

Class base_Data_Array
	Private pArray

	Private Sub Class_Initialize()

	End Sub


	' Properties:


	Public Property Get Allocated()
		If IsArrayAllocated(pArray) Then
			Allocated = True
		Else
			Allocated = False
		End If
	End Property

	Public Default Property Get Item(intIndex)
		If IsArrayAllocated(pArray) Then
    			If IsObject(pArray(intIndex)) Then
        			Set Item = pArray(intIndex)
    			Else
        			Item = pArray(intIndex)
    			End If
		End If
	End Property

	Public Property Let Item(intIndex, varInput)
		If IsArrayAllocated(pArray) Then pArray(intIndex) = varInput
	End Property

	Public Property Set Item(intIndex, objInput)
		If IsArrayAllocated(pArray) Then Set pArray(intIndex) = varInput
	End Property

	Public Property Get Length()
		If IsArrayAllocated(pArray) Then Length = UBound(pArray) + 1
	End Property


	' Methods:


	Public Sub Append(varInput)
		If IsArray(varInput) Then
			Dim i

			Extend UBound(varInput) + 1, True

			For i = 0 To UBound(varInput)
				If IsObject(varInput(i)) Then
					Set pArray(UBound(pArray) - UBound(varInput) + i) = varInput(i)
				Else
					pArray(UBound(pArray) - UBound(varInput) + i) = varInput(i)
				End If
			Next
		Else
			Extend 1, True

			If IsObject(varInput) Then
				Set pArray(UBound(pArray)) = varInput
			Else
				pArray(UBound(pArray)) = varInput
			End If
		End If
	End Sub

	Public Sub Insert(varInput, intIndex)
		Dim objArray
		Set objArray = Slice(LBound(pArray), intIndex - 1)
		objArray.Append varInput
		If UBound(pArray) > intIndex Then objArray.Append Slice(intIndex, UBound(pArray)).ToArray()
		Me.FromArray objArray.ToArray()
		Set objArray = Nothing
	End Sub

	Public Sub RemoveAt(intIndex)
		If IsArrayAllocated(pArray) Then
			Dim objArray, _
				i

			Set objArray = New v_Data_Array

			For i = 0 To UBound(pArray)
  				Do
    					If i = intIndex Then Exit Do
        				
					objArray.Append pArray(i)
  				Loop While False
			Next

			Me.FromArray objArray.ToArray()
			Set objArray = Nothing
		End If
	End Sub

	Public Sub RemoveValue(varInput)
		RemoveAt IndexOf(varInput)
	End Sub

	Public Sub Push(varInput)
		Append(varInput)
	End Sub

	Public Function Pop(intIndex)
		Pop = Item(intIndex)
		RemoveAt intIndex
	End Function
	
	Public Function Slice(intStart, intEnd)
		Dim objArr, _
			i

		Set objArr = New v_Data_Array

		For i = intStart To intEnd
			objArr.Append Me(i)
		Next

		Set Slice = objArr
	End Function

	Public Sub Splice(intIndex, intRemove, arrInput)
		Dim objArray
		Set objArray = Slice(LBound(pArray), intIndex - 1)
		objArray.Append arrInput
		If UBound(pArray) > (intIndex + intRemove) Then objArray.Append Slice(intIndex + intRemove, UBound(pArray)).ToArray()
		Me.FromArray objArray.ToArray()
		Set objArray = Nothing
	End Sub

	Public Sub Extend(intSize, blnPreserve)
		If IsArrayAllocated(pArray) Then
			If blnPreserve Then
				ReDim Preserve pArray(UBound(pArray) + intSize)
			Else
				ReDim pArray(UBound(pArray) + intSize)
			End If
		Else
			pArray = Array()
			ReDim pArray(intSize - 1)
		End If
	End Sub

	Public Sub Resize(intSize, blnPreserve)
		If IsArrayAllocated(pArray) Then
			If blnPreserve Then
				ReDim Preserve pArray(intSize - 1)
			Else
				ReDim pArray(intSize - 1)
			End If
		Else
			pArray = Array()
			ReDim pArray(intSize - 1)
		End If
	End Sub

	Public Function IndexOf(varInput)
		Dim intIndex, _
			i

		For i = 0 To UBound(pArray)
			If IsObject(varInput) And IsObject(pArray(i)) Then
				If pArray(i) Is varInput Then intIndex = i
			ElseIf Not IsObject(pArray(i)) And Not IsObject(varInput) Then
				If pArray(i) = varInput Then intIndex = i
			End If
			If Not IsEmpty(intIndex) Then Exit For
		Next

		If Not IsEmpty(intIndex) Then IndexOf = intIndex
	End Function

	Public Sub Sort()
		QuickSort pArray, LBound(pArray), UBound(pArray)
	End Sub

	Public Sub Reverse()
		If IsArrayAllocated(pArray) Then
			Dim objArray, _
				i

			Set objArray = New v_Data_Array

			For i = UBound(pArray) To 0 Step -1
				objArray.Append pArray(i)
			Next

			Me.FromArray objArray.ToArray()
			Set objArray = Nothing
		End If
	End Sub

	Public Sub FromArray(arrInput)
		If IsArray(arrInput) Then
			If IsArrayAllocated(pArray) Then Clear
			pArray = arrInput
		End If
	End Sub

	Public Function ToArray()
		If IsArrayAllocated(pArray) Then ToArray = pArray
	End Function

	Public Sub FromString(strInput, varDelimiter)
		If TypeName(strInput) = "String" Then
			If IsArrayAllocated(pArray) Then Clear
			pArray = Split(strInput, CStr(varDelimiter))
		End If
	End Sub

	Public Function ToString()
		If IsArrayAllocated(pArray) Then ToString = Join(pArray, ",")
	End Function

	Public Function Clone()
		Dim objArray
		Set objArray = New base_Data_Array
		objArray.FromArray Me.ToArray()
		Set Clone = objArray
	End Function

	Public Sub Clear()
		If IsArrayAllocated(pArray) Then Erase pArray
	End Sub

	Private Sub Class_Terminate()
		Clear()
	End Sub
End Class

If WScript.ScriptName = "base_Data_Array.vbs" Then
	Dim objArray, _
		i

	Set objArray = New base_Data_Array

	' objArray.FromArray Array("Banana", "Orange", "Lemon", "Apple", "Mango")
	' objArray.Append Array("Car", "Bus", "Train", "Boat")
	' objArray.Splice 2, 0, Array("Lime", "Kiwi")

	objArray.FromString "Banana, Apple, Mango, Kiwi", ","

	WScript.Echo objArray(0)

	' For i = 0 To objArray.Length - 1
	' 	WScript.Echo "objArray(" & i & ") = " & objArray(i)
	' Next
End If
