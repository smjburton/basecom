VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "base_Data_Collection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_Collection() As Variant

Private Sub Class_Initialize()

End Sub


' Properties


Public Property Get Count() As Integer
    If IsArrayAllocated(p_Collection) Then Count = UBound(p_Collection) + 1
End Property

' *** Make default property
Public Property Get Item( _
    ByVal intIndex As Integer _
    ) As Variant
    
    If IsArrayAllocated(p_Collection) Then
            If IsObject(p_Collection(intIndex)) Then
                Set Item = p_Collection(intIndex)
            Else
                Item = p_Collection(intIndex)
            End If
    End If
End Property


' Methods


Public Sub Add( _
    ByVal varItem As Variant _
    )
    
    Extend 1

    If IsObject(varItem) Then
        Set p_Collection(UBound(p_Collection)) = varItem
    Else
        p_Collection(UBound(p_Collection)) = varItem
    End If
End Sub

Public Sub Clear()
    If IsArrayAllocated(p_Collection) Then Erase p_Collection
End Sub

Public Function Contains( _
    ByVal varItem As Variant _
    ) As Boolean
    
    Dim blnContains As Boolean, _
        intIndex As Integer

    blnContains = False

    For intIndex = 0 To UBound(p_Collection)
        If IsObject(varItem) And IsObject(p_Collection(intIndex)) Then
            If p_Collection(intIndex) Is varItem Then blnContains = True
        ElseIf Not IsObject(p_Collection(intIndex)) And Not IsObject(varItem) Then
            If p_Collection(intIndex) = varItem Then blnContains = True
        End If
        If blnContains Then Exit For
    Next intIndex

    Contains = blnContains
End Function

Public Sub Remove( _
    ByVal intIndex As Integer _
    )
    
    If IsArrayAllocated(p_Collection) Then
        Dim objCollection As base_Data_Collection, _
            i As Integer

        Set objCollection = New base_Data_Collection

        For i = 0 To UBound(pCollection)
            Do
                If i = intIndex Then Exit Do
                    
                objCollection.Add p_Collection(i)
            Loop While False
        Next

        Me.Clear

        For i = 0 To objCollection.Count - 1
            Me.Add objCollection(i)
        Next
    End If
End Sub


' Helper Methods


Private Sub Extend( _
    ByVal intSize As Integer _
    )
    
    If IsArrayAllocated(p_Collection) Then
        ReDim Preserve p_Collection(UBound(p_Collection) + intSize)
    Else
        p_Collection = Array()
        ReDim p_Collection(intSize - 1)
    End If
End Sub

Private Function IsArrayAllocated( _
    ByRef arrArray() As Variant _
    ) As Boolean
    
    IsArrayAllocated = False
    
    If IsArray(arrArray) Then
        On Error Resume Next
        
        Dim intUpperBound As Integer
        intUpperBound = UBound(arrArray)
        If (Err.Number = 0) And (intUpperBound >= 0) Then IsArrayAllocated = True
    End If
End Function

Private Sub Class_Terminate()

End Sub
