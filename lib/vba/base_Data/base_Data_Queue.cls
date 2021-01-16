VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "base_Data_Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_Queue As Object

Private Sub Class_Initialize()
    Set p_Queue = CreateObject("System.Collections.Queue")
End Sub


' Properties


Public Property Get Count() As Integer
    Count = p_Queue.Count
End Property

Public Property Get IsSynchronized() As Boolean
    IsSynchronized = p_Queue.IsSynchronized
End Property

Public Property Get SyncRoot() As Variant
    SyncRoot = p_Queue.SyncRoot
End Property


' Methods


Public Sub Clear()
    p_Queue.Clear
End Sub

Public Function Clone() As Object
    Set Clone = p_Queue.Clone()
End Function

Public Function Contains( _
    ByVal varInput As Variant _
    ) As Boolean
    
    Contains = p_Queue.Contains(varInput)
End Function

Public Function Dequeue() As Variant
    If IsObject(p_Queue.Peek()) Then
        Set Dequeue = p_Queue.Dequeue()
    Else
        Dequeue = p_Queue.Dequeue()
    End If
End Function

Public Sub Enqueue( _
    ByVal varInput As Variant _
    )
    
    p_Queue.Enqueue varInput
End Sub

Public Function Equals( _
    ByVal varInput As Variant _
    ) As Boolean
     
    Equals = p_Queue.Equals(varInput)
End Function

Public Function GetEnumerator() As Object
    Set GetEnumerator = p_Queue.GetEnumerator()
End Function

Public Function GetHashCode() As String
    GetHashCode = p_Queue.GetHashCode()
End Function

Public Function GetType() As Object
    Set GetType = p_Queue.GetType()
End Function

Public Function Peek() As Variant
    If IsObject(p_Queue.Peek()) Then
        Set Peek = p_Queue.Peek()
    Else
        Peek = p_Queue.Peek()
    End If
End Function

Public Function ToArray() As Variant()
     ToArray = p_Queue.ToArray()
End Function

Public Function ToString() As String
     ToString = p_Queue.ToString()
End Function

Public Sub TrimToSize()
     p_Queue.TrimToSize
End Sub

Private Sub Class_Terminate()
    Set p_Queue = Nothing
End Sub
