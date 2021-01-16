VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "base_Data_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_List As Object

Private Sub Class_Initialize()
    Set p_List = CreateObject("System.Collections.SortedList")
End Sub


' Properties


Public Property Get Capacity() As Integer
    Capacity = p_List.Capacity
End Property

Public Property Let Capacity( _
    ByVal intCapacity As Integer _
    )
    
    p_List.Capacity = intCapacity
End Property

Public Property Get Count() As Integer
    Count = p_List.Count
End Property

Public Property Get IsFixedSize() As Boolean
    IsFixedSize = p_List.IsFixedSize
End Property

Public Property Get IsReadOnly() As Boolean
    IsReadOnly = p_List.IsReadOnly
End Property

Public Property Get IsSynchronized() As Boolean
    IsSynchronized = p_List.IsSynchronized
End Property

' *** Make Default property
Public Property Get Item( _
    ByVal strKey As String _
    ) As Variant
    
    If IsObject(p_List.Item(strKey)) Then
        Set Item = p_List.Item(strKey)
    Else
        Item = p_List.Item(strKey)
    End If
End Property

Public Property Let Item( _
    ByVal strKey As String, _
    ByVal varItem As Variant _
    )
    
    pList.Item(strKey) = varItem
End Property

Public Property Set Item( _
    ByVal strKey As String, _
    ByVal varItem As Variant _
    )
    
    Set pList.Item(strKey) = varItem
End Property

Public Property Get Keys() As Variant()
    Set Keys = p_List.Keys
End Property

Public Property Get SyncRoot() As Variant
    SyncRoot = p_List.SyncRoot
End Property

Public Property Get Values() As Variant()
    Set Values = p_List.Values
End Property


' Methods


Public Sub Add( _
    ByVal strKey As String, _
    ByVal varItem As Variant _
    )
    
    p_List.Add strKey, varItem
End Sub

Public Sub Clear()
    p_List.Clear
End Sub

Public Function Clone() As Object
    Set Clone = p_List.Clone()
End Function

Public Function Contains( _
    ByVal varItem As Variant _
    ) As Boolean
    
    Contains = p_List.Contains(varItem)
End Function

Public Function ContainsKey( _
    ByVal strKey As String _
    ) As Boolean
    
    ContainsKey = p_List.ContainsKey(strKey)
End Function

Public Function ContainsValue( _
    ByVal varItem As Variant _
    ) As Boolean
    
    ContainsValue = p_List.ContainsValue(varItem)
End Function

Public Function Equals( _
    ByVal objItem As Object _
    ) As Boolean
    
    Equals = p_List.Equals(objItem)
End Function

Public Function GetByIndex( _
    ByVal intIndex As Integer _
    ) As Variant
    
    If IsObject(pList.GetByIndex(intIndex)) Then
        Set GetByIndex = p_List.GetByIndex(intIndex)
    Else
        GetByIndex = p_List.GetByIndex(intIndex)
    End If
End Function

Public Function GetEnumerator() As Object
    Set GetEnumerator = p_List.GetEnumerator()
End Function

Public Function GetHashCode() As String
    GetHashCode = p_List.GetHashCode()
End Function

Public Function GetKey( _
    ByVal intIndex As Integer _
    ) As String
    
    GetKey = p_List.GetKey(intIndex)
End Function

Public Function GetKeyList() As Object
    Set GetKeyList = p_List.GetKeyList()
End Function

Public Function GetType() As Object
    Set GetType = p_List.GetType()
End Function

Public Function GetValueList() As Object
    Set GetValueList = p_List.GetValueList()
End Function

Public Function IndexOfKey( _
    ByVal strKey As String _
    ) As Integer
    
    IndexOfKey = p_List.IndexOfKey(strKey)
End Function

Public Function IndexOfValue( _
    ByVal varItem As Variant _
    ) As Integer
    
    IndexOfValue = p_List.IndexOfValue(varItem)
End Function

Public Sub Remove( _
    ByVal varItem As Variant _
    )
    
    p_List.Remove varItem
End Sub

Public Sub RemoveAt( _
    ByVal intIndex As Integer _
    )
    
    p_List.RemoveAt intIndex
End Sub

Public Sub SetByIndex( _
    ByVal intIndex As Integer, _
    ByVal varItem As Variant _
    )
    
    p_List.SetByIndex intIndex, varItem
End Sub

Public Function ToString() As String
    ToString = p_List.ToString()
End Function

Public Sub TrimToSize()
    p_List.TrimToSize
End Sub

Private Sub Class_Terminate()
    Set p_List = Nothing
End Sub
