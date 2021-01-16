VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "base_Data_Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_Stack As Object

Private Sub Class_Initialize()
    Set p_Stack = CreateObject("System.Collections.Stack")
End Sub


' Properties


Public Property Get Count() As Integer
    Count = p_Stack.Count
End Property

Public Property Get IsSynchronized() As Boolean
    IsSynchronized = p_Stack.IsSynchronized
End Property

Public Property Get SyncRoot() As Variant
    SyncRoot = p_Stack.SyncRoot
End Property


' Methods


Public Sub Clear()
    p_Stack.Clear
End Sub

Public Function Clone() As Object
    Set Clone = p_Stack.Clone()
End Function

Public Function Contains( _
    ByVal varInput As Variant _
    ) As Boolean
    
    Contains = p_Stack.Contains(varInput)
End Function

Public Function Equals( _
    ByVal varInput As Variant _
    ) As Boolean
    
    Equals = p_Stack.Equals(varInput)
End Function

Public Function GetEnumerator() As Object
    Set GetEnumerator = p_Stack.GetEnumerator()
End Function

Public Function GetHashCode() As String
    GetHashCode = p_Stack.GetHashCode()
End Function

Public Function GetType() As Object
    Set GetType = p_Stack.GetType()
End Function

Public Function Peek() As Variant
    If IsObject(p_Stack.Peek()) Then
        Set Peek = p_Stack.Peek()
    Else
        Peek = p_Stack.Peek()
    End If
End Function

Public Function Pop() As Variant
    If IsObject(p_Stack.Peek()) Then
        Set Pop = p_Stack.Pop()
    Else
        Pop = p_Stack.Pop()
    End If
End Function

Public Sub Push( _
    ByVal varInput As Variant _
    )
    
    p_Stack.Push varInput
End Sub

Public Function Synchronized( _
    ByVal objStack As Object _
    ) As Object
    
    Set Synchronized = p_Stack.Synchronized(objStack)
End Function

Public Function ToArray() As Variant()
    ToArray = p_Stack.ToArray()
End Function

Public Function ToString() As String
    ToString = p_Stack.ToString()
End Function

Private Sub Class_Terminate()
    Set p_Stack = Nothing
End Sub
