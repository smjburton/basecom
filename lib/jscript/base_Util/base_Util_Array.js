Option Explicit

Function IsArrayAllocated(arrArray)
	IsArrayAllocated = False
	If IsArray(arrArray) Then
		On Error Resume Next
		Dim ub : ub = UBound(arrArray)
		If (Err.Number = 0) And (ub >= 0) Then IsArrayAllocated = True
	End If  
End Function

Sub BubbleSort()

End Sub

Sub QuickSort(ByRef arrArray, intLoBound, intHiBound)
	Dim varPivot, _
		intLoSwap, _
		intHiSwap, _
		varTemp

	If intHiBound - intLoBound = 1 Then
		If arrArray(intLoBound) > arrArray(intHiBound) Then
			varTemp = arrArray(intLoBound)
			arrArray(intLoBound) = arrArray(intHiBound)
			arrArray(intHiBound) = varTemp
		End If
	End If

	varPivot = arrArray(CInt((intLoBound + intHiBound) / 2))
	arrArray(CInt((intLoBound + intHiBound) / 2)) = arrArray(intLoBound)
	arrArray(intLoBound) = varPivot
	intLoSwap = intLoBound + 1
	intHiSwap = intHiBound
  
	Do
		While intLoSwap < intHiSwap and arrArray(intLoSwap) <= varPivot
			intLoSwap = intLoSwap + 1
		Wend

		While arrArray(intHiSwap) > varPivot
			intHiSwap = intHiSwap - 1
		Wend

		If intLoSwap < intHiSwap Then
			varTemp = arrArray(intLoSwap)
			arrArray(intLoSwap) = arrArray(intHiSwap)
			arrArray(intHiSwap) = varTemp
		End If
	Loop While intLoSwap < intHiSwap
  
	arrArray(intLoBound) = arrArray(intHiSwap)
	arrArray(intHiSwap) = varPivot
  
	If intLoBound < (intHiSwap - 1) Then Call QuickSort(arrArray, intLoBound, intHiSwap - 1)
	If intHiSwap + 1 < intHiBound Then Call QuickSort(arrArray, intHiSwap + 1, intHiBound)
End Sub

' Sub Reverse(ByRef arrArray)
' 	Dim i, j, idxLast, idxHalf, strHolder
' 
' 	idxLast = UBound( myArray )
' 	idxHalf = Int( idxLast / 2 )
' 
' 	For i = 0 To idxHalf
' 		strHolder              = myArray( i )
' 		myArray( i )           = myArray( idxLast - i )
' 		myArray( idxLast - i ) = strHolder
' 	Next
' End Sub

' function reverse(array) {
'  var first = null;
'  var last = null;
'  var tmp = null;
'  var length = array.length;
'
'  for (first = 0, last = length - 1; first < length / 2; first++, last--) {
'    tmp = array[first];
'    array[first] = array[last];
'    array[last] = tmp;
'  }
' }

If WScript.ScriptName = "v_Util_Array.vbs" Then

End If