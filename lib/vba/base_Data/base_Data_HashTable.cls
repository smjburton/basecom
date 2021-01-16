VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "base_Data_HashTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_HashTable As Object

Private Sub Class_Initialize()
    Set p_HashTable = CreateObject("System.Collections.Hashtable")
End Sub


' Properties


Public Property Get Count() As Integer
    Count = p_HashTable.Count
End Property

Public Property Get IsFixedSize() As Boolean
    IsFixedSize = p_HashTable.IsFixedSize
End Property

Public Property Get IsReadOnly() As Boolean
    IsReadOnly = p_HashTable.IsReadOnly
End Property

Public Property Get IsSynchronized() As Boolean
    IsSynchronized = p_HashTable.IsSynchronized
End Property

' *** Make default property
Public Property Get Item( _
    ByVal strKey As String _
    ) As Variant
    
    If IsObject(p_HashTable(strKey)) Then
        Set Item = p_HashTable(strKey)
    Else
        Item = p_HashTable.Item(strKey)
    End If
End Property

Public Property Let Item( _
    ByVal strKey As String, _
    ByVal varInput As Variant _
    )
    
    p_HashTable(strKey) = varInput
End Property

Public Property Set Item( _
    ByVal strKey As String, _
    ByVal varInput As Variant _
    )
    
    Set p_HashTable(strKey) = varInput
End Property

Public Property Get Keys() As Variant()
    Set Keys = p_HashTable.Keys
End Property

Public Property Get SyncRoot() As Variant
    SyncRoot = p_HashTable.SyncRoot
End Property

Public Property Get Values() As Variant()
    Set Values = p_HashTable.Values
End Property


' Methods


Public Sub Add( _
    ByVal strKey As String, _
    ByVal varItem As Variant _
    )
    
    p_HashTable.Add strKey, varItem
End Sub

Public Sub Clear()
    p_HashTable.Clear
End Sub

Public Function Clone() As Object
    Set Clone = p_HashTable.Clone()
End Function

Public Function Contains( _
    ByVal strKey As String _
    ) As Boolean
    
    Contains = p_HashTable.Contains(strKey)
End Function

Public Function ContainsKey( _
    ByVal strKey As String _
    ) As Boolean
    
    ContainsKey = p_HashTable.ContainsKey(strKey)
End Function

Public Function ContainsValue( _
    ByVal varItem As Variant _
    ) As Boolean
    
    ContainsValue = p_HashTable.ContainsValue(varItem)
End Function

Public Function Equals( _
    ByVal objItem As Object _
    ) As Boolean
    
    Equals = p_HashTable.Equals(objItem)
End Function

Public Function GetEnumerator() As Variant
    Set GetEnumerator = p_HashTable.GetEnumerator()
End Function

Public Function GetHashCode() As Variant
    GetHashCode = p_HashTable.GetHashCode()
End Function

Public Function GetType() As Variant
    GetType = p_HashTable.GetType()
End Function

Public Function Remove( _
    ByVal strKey As String _
    )
    
    p_HashTable.Remove strKey
End Function

Public Function ToString() As String
    ToString = p_HashTable.ToString()
End Function

Private Sub Class_Terminate()
    Set p_HashTable = Nothing
End Sub
