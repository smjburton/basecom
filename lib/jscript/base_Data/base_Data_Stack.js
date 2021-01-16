Class v_Data_Stack
	Private pStack

	Private Sub Class_Initialize()
		Set pStack = CreateObject("System.Collections.Stack")
	End Sub


	' Properties 


	Public Property Get Count()
		Count = pStack.Count
	End Property

	Public Property Get IsSynchronized()
		IsSynchronized = pStack.IsSynchronized
	End Property

	Public Property Get SyncRoot()
		SyncRoot = pStack.SyncRoot
	End Property

	
	' Methods


	Public Sub Clear()
		pStack.Clear()
	End Sub

	Public Function Clone()
		Set Clone = pStack.Clone()
	End Function

	Public Function Contains(objInput)
		Contains = pStack.Contains(objInput)
	End Function

	Public Function Equals(objInput)
		Equals = pStack.Equals(objInput)
	End Function

	Public Function GetEnumerator()
		Set GetEnumerator = pStack.GetEnumerator()
	End Function

	Public Function GetHashCode()
		GetHashCode = pStack.GetHashCode()
	End Function

	Public Function GetType()
		Set GetType = pStack.GetType()
	End Function

	Public Function Peek()
		If IsObject(pStack.Peek()) Then
			Set Peek = pStack.Peek()
		Else
			Peek = pStack.Peek()
		End If
	End Function

	Public Function Pop()
		If IsObject(pStack.Peek()) Then
			Set Pop = pStack.Pop()
		Else
			Pop = pStack.Pop()
		End If
	End Function

	Public Sub Push(objInput)
		pStack.Push objInput
	End Sub

	Public Function Synchronized(objStack)
		Set Synchronized = pStack.Synchronized(objStack)
	End Function

	Public Function ToArray()
		ToArray = pStack.ToArray()
	End Function

	Public Function ToString()
		ToString = pStack.ToString()
	End Function

	Private Sub Class_Terminate()
		Set pStack = Nothing
	End Sub
End Class

If WScript.ScriptName = "v_Data_Stack.vbs" Then
	Dim stack

	Set stack = New v_Data_Stack

	stack.Push "Apple"
	stack.Push "Orange"
	stack.Push "Banana"
	stack.Push "Strawberry"

	WScript.Echo stack.SyncRoot
End If