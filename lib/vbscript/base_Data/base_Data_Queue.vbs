Option Explicit

Class base_Data_Queue
	Private pQueue

	Private Sub Class_Initialize()
		Set pQueue = CreateObject("System.Collections.Queue")
	End Sub


	' Properties


	Public Property Get Count()
		Count = pQueue.Count
	End Property

	Public Property Get IsSynchronized()
		IsSynchronized = pQueue.IsSynchronized
	End Property

	Public Property Get SyncRoot()
		SyncRoot = pQueue.SyncRoot
	End Property


	' Methods


	Public Sub Clear()
		 pQueue.Clear()
	End Sub

	Public Function Clone()
		 Set Clone = pQueue.Clone()
	End Function

	Public Function Contains(objInput)
		 Contains = pQueue.Contains(objInput)
	End Function

	Public Function Dequeue()
		If IsObject(pQueue.Peek()) Then
			Set Dequeue = pQueue.Dequeue()
		Else
			Dequeue = pQueue.Dequeue()
		End If
	End Function

	Public Sub Enqueue(objInput)
		 pQueue.Enqueue(objInput)
	End Sub

	Public Function Equals(objInput)
		 Equals = pQueue.Equals(objInput)
	End Function

	Public Function GetEnumerator()
		 Set GetEnumerator = pQueue.GetEnumerator()
	End Function

	Public Function GetHashCode()
		 GetHashCode = pQueue.GetHashCode()
	End Function

	Public Function GetType()
		 Set GetType = pQueue.GetType()
	End Function

	Public Function Peek()
		If IsObject(pQueue.Peek()) Then
			Set Peek = pQueue.Peek()
		Else
			Peek = pQueue.Peek()
		End If
	End Function

	Public Function ToArray()
		 ToArray = pQueue.ToArray()
	End Function

	Public Function ToString()
		 ToString = pQueue.ToString()
	End Function

	Public Sub TrimToSize()
		 pQueue.TrimToSize()
	End Sub

	Private Sub Class_Terminate()
		Set pQueue = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Data_Queue.vbs" Then
	Dim queue
	Set queue = New base_Data_Queue

	queue.Enqueue "Dog"
	queue.Enqueue "Cat"
	queue.Enqueue "Bird"
	queue.Enqueue "Lizard"

	WScript.Echo queue.Dequeue()
End If
