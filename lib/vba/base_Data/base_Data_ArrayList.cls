VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "base_Data_ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_ArrayList as Object

Private Sub Class_Initialize()
    Set p_ArrayList = CreateObject("System.Collections.ArrayList")
End Sub


' Properties


Public Property Get Capacity() As Integer
    Capacity = p_ArrayList.Capacity
End Property

Public Property Let Capacity( _
    ByVal intSize As Integer _
    )
    
    p_ArrayList.Capacity = intSize
End Property

Public Property Get Count() As Variant
    Count = p_ArrayList.Count
End Property

Public Property Get IsFixedSize() As Boolean
    IsFixedSize = p_ArrayList.IsFixedSize
End Property

Public Property Get IsReadOnly() As Boolean
    IsReadOnly = p_ArrayList.IsReadOnly
End Property

Public Property Get IsSynchronized() As Boolean
    IsSynchronized = p_ArrayList.IsSynchronized
End Property

Public Property Get Item( _
    ByVal intIndex As Integer _
    ) As Variant
    
    Attribute Value.VB_UserMemId = 0

    If IsObject(p_ArrayList(intIndex)) Then
        Set Item = p_ArrayList(intIndex)
    Else
        Item = p_ArrayList(intIndex)
    End If
End Property

Public Property Let Item( _
    ByVal intIndex As Integer, _
    ByVal varInput As Variant _
    )
    
    p_ArrayList(intIndex) = varInput
End Property

Public Property Set Item( _
    ByVal intIndex As Integer, _
    ByVal varInput As Variant _
    )
    
    Set p_ArrayList(intIndex) = varInput
End Property

Public Property Get SyncRoot() As Variant
    SyncRoot = p_ArrayList.SyncRoot
End Property


' Methods


Public Sub Add( _
    ByVal varItem As Variant _
    )
    
    p_ArrayList.Add varItem
End Sub

Public Sub Clear()
    p_ArrayList.Clear
End Sub

Public Function Clone()
    Set Clone = p_ArrayList.Clone()
End Function

Public Function Contains( _
    ByVal varItem As Variant _
    ) As Boolean
    
    Contains = p_ArrayList.Contains(varItem)
End Function

Public Function Equals( _
    ByVal objItem As Object _
    ) As Boolean
    Equals = p_ArrayList.Equals(objItem)
End Function

Public Function GetEnumerator( _
    ByVal intStart As Integer, _
    ByVal intEnd As Integer _
    ) As Variant
    
    Set GetEnumerator = p_ArrayList.GetEnumerator(intStart, intEnd)
End Function

Public Function GetHashCode() As Variant
    GetHashCode = p_ArrayList.GetHashCode()
End Function

Public Function GetType() As Variant
    GetType = p_ArrayList.GetType()
End Function

Public Sub Insert( _
    ByVal intIndex As Integer, _
    ByVal varItem As Variant _
    )
    
    p_ArrayList.Insert intIndex, varItem
End Sub

Public Sub Remove( _
    ByVal varItem As Variant _
    )
    
    p_ArrayList.Remove varItem
End Sub

Public Sub RemoveAt( _
    ByVal intIndex As Integer _
    )
    
    p_ArrayList.RemoveAt intIndex
End Sub

Public Sub Reverse()
    p_ArrayList.Reverse
End Sub

Public Sub Sort()
    p_ArrayList.Sort
End Sub

Public Function ToArray() As Variant()
    ToArray = p_ArrayList.ToArray()
End Function

Public Function ToString() As String
    ToString = p_ArrayList.ToString()
End Function

Public Function TrimToSize() As Variant
    TrimToSize = p_ArrayList.TrimToSize()
End Function

Private Sub Class_Terminate()
    Set p_ArrayList = Nothing
End Sub
