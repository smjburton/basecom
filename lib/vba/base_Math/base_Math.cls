Option Explicit

Function Mean(arrValues)
	If TypeName(arrValues) = "Variant()" Then
		Dim intSum, _
			i

		intSum = 0

		For i = 0 To UBound(arrValues)
			intSum = intSum + arrValues(i)
		Next

		Mean = intSum / (UBound(arrValues) + 1)
	End If
End Function

Function StdDev(arrValues)
	If TypeName(arrValues) = "Variant()" Then
		Dim intAverage, _
			intSumSq, _
			i

		intAverage = Mean(arrValues)
		intSumSq = 0

		For i = 0 To UBound(arrValues)
			intSumSq = intSumSq + (arrValues(i) - intAverage) ^ 2
		Next

		StdDev = Sqr(intSumSq / UBound(arrValues))
	End If
End Function
