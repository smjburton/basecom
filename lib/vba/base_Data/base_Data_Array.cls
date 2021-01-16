VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "base_Data_Array"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_Array() As Variant

Private Sub Class_Initialize()

End Sub

' Properties:

Public Property Get Allocated() As Boolean
    If IsArrayAllocated(p_Array) Then
        Allocated = True
    Else
        Allocated = False
    End If
End Property

Public Property Get Item( _
    ByVal intIndex As Integer _
    ) As Variant
    
    If IsArrayAllocated(p_Array) Then
            If IsObject(p_Array(intIndex)) Then
                Set Item = p_Array(intIndex)
            Else
                Item = p_Array(intIndex)
            End If
    End If
End Property

Public Property Let Item( _
    ByVal intIndex As Integer, _
    ByVal varInput As Variant _
    )
    
    If IsArrayAllocated(p_Array) Then p_Array(intIndex) = varInput
End Property

Public Property Set Item( _
    ByVal intIndex As Integer, _
    ByVal objInput As Object _
    )
    
    If IsArrayAllocated(p_Array) Then Set p_Array(intIndex) = varInput
End Property

Public Property Get Length() As Integer
    If IsArrayAllocated(p_Array) Then Length = UBound(p_Array) + 1
End Property

' Methods:

Public Sub Append( _
    ByVal varInput As Variant _
    )
    
    If IsArray(varInput) Then
        Dim intIndex As Integer

        Extend UBound(varInput) + 1, True

        For intIndex = 0 To UBound(varInput)
            If IsObject(varInput(intIndex)) Then
                Set p_Array(UBound(p_Array) - UBound(varInput) + intIndex) = varInput(intIndex)
            Else
                p_Array(UBound(p_Array) - UBound(varInput) + intIndex) = varInput(intIndex)
            End If
        Next intIndex
    Else
        Extend 1, True

        If IsObject(varInput) Then
            Set p_Array(UBound(p_Array)) = varInput
        Else
            p_Array(UBound(p_Array)) = varInput
        End If
    End If
End Sub

Public Sub Insert( _
    ByVal varInput As Variant, _
    ByVal intIndex As Integer _
    )
    
    Dim objArray As Object
    Set objArray = Slice(LBound(p_Array), intIndex - 1)
    objArray.Append varInput
    If UBound(p_Array) > intIndex Then objArray.Append Slice(intIndex, UBound(p_Array)).ToArray()
    Me.FromArray objArray.ToArray()
    Set objArray = Nothing
End Sub

Public Sub RemoveAt( _
    ByVal intIndex As Integer _
    )
    
    If IsArrayAllocated(p_Array) Then
        Dim objArray As Object, _
            intArrIndex

        Set objArray = New v_Data_Array

        For intArrIndex = 0 To UBound(p_Array)
            Do
                If intArrIndex = intIndex Then Exit Do
                objArray.Append p_Array(intArrIndex)
            Loop While False
        Next

        Me.FromArray objArray.ToArray()
        Set objArray = Nothing
    End If
End Sub

Public Sub RemoveValue( _
    ByVal varInput As Variant _
    )
    
    RemoveAt IndexOf(varInput)
End Sub

Public Sub Push( _
    ByVal varInput As Variant _
    )
    
    Append varInput
End Sub

Public Function Pop( _
    ByVal intIndex As Integer _
    ) As Variant
    
    Pop = Item(intIndex)
    RemoveAt intIndex
End Function

Public Function Slice( _
    ByVal intStart As Integer, _
    ByVal intEnd As Integer _
    ) As Object
    
    Dim objArr As Object, _
        intIndex As Integer

    Set objArr = New base_Data_Array

    For intIndex = intStart To intEnd
        objArr.Append Me.Item(intIndex)
    Next

    Set Slice = objArr
End Function

Public Sub Splice( _
    ByVal intIndex As Integer, _
    ByVal intRemove As Integer, _
    ByRef arrInput() As Variant _
    )
    
    Dim objArray As Object
    Set objArray = Slice(LBound(p_Array), intIndex - 1)
    objArray.Append arrInput
    If UBound(p_Array) > (intIndex + intRemove) Then objArray.Append Slice(intIndex + intRemove, UBound(p_Array)).ToArray()
    Me.FromArray objArray.ToArray()
    Set objArray = Nothing
End Sub

Public Sub Extend( _
    ByVal intSize As Integer, _
    ByVal blnPreserve As Boolean _
    )
    
    If IsArrayAllocated(p_Array) Then
        If blnPreserve Then
            ReDim Preserve p_Array(UBound(p_Array) + intSize)
        Else
            ReDim p_Array(UBound(p_Array) + intSize)
        End If
    Else
        p_Array = Array()
        ReDim p_Array(intSize - 1)
    End If
End Sub

Public Sub Resize( _
    ByVal intSize As Integer, _
    ByVal blnPreserve As Boolean _
    )
    
    If IsArrayAllocated(p_Array) Then
        If blnPreserve Then
            ReDim Preserve p_Array(intSize - 1)
        Else
            ReDim p_Array(intSize - 1)
        End If
    Else
        p_Array = Array()
        ReDim p_Array(intSize - 1)
    End If
End Sub

Public Function IndexOf( _
    ByVal varInput As Variant _
    ) As Integer
    
    Dim intArrIndex As Integer, _
        intIndex As Integer

    For intIndex = 0 To UBound(p_Array)
        If IsObject(varInput) And IsObject(p_Array(intIndex)) Then
            If p_Array(intIndex) Is varInput Then intArrIndex = intIndex
        ElseIf Not IsObject(p_Array(intIndex)) And Not IsObject(varInput) Then
            If p_Array(intIndex) = varInput Then intArrIndex = intIndex
        End If
        If Not IsEmpty(intArrIndex) Then Exit For
    Next

    If Not IsEmpty(intArrIndex) Then IndexOf = intArrIndex
End Function

Public Sub Sort()
    QuickSort p_Array, LBound(p_Array), UBound(p_Array)
End Sub

Public Sub Reverse()
    If IsArrayAllocated(p_Array) Then
        Dim objArray As Object, _
            intIndex As Integer

        Set objArray = New v_Data_Array

        For intIndex = UBound(p_Array) To 0 Step -1
            objArray.Append p_Array(intIndex)
        Next

        Me.FromArray objArray.ToArray()
        Set objArray = Nothing
    End If
End Sub

Public Sub FromArray( _
    ByRef arrInput() As Variant _
    )
    
    If IsArray(arrInput) Then
        If IsArrayAllocated(p_Array) Then Me.Clear
        p_Array = arrInput
    End If
End Sub

Public Function ToArray() As Variant
    If IsArrayAllocated(p_Array) Then ToArray = p_Array
End Function

Public Sub FromString( _
    ByVal strInput As String, _
    ByVal varDelimiter As Variant _
    )
    
    If TypeName(strInput) = "String" Then
        If IsArrayAllocated(p_Array) Then Me.Clear
        p_Array = Split(strInput, CStr(varDelimiter))
    End If
End Sub

Public Function ToString() As String
    If IsArrayAllocated(p_Array) Then ToString = Join(p_Array, ", ")
End Function

Public Function Clone() As Object
    Dim objArray As Object
    Set objArray = New v_Data_Array
    objArray.FromArray Me.ToArray()
    Set Clone = objArray
End Function

Public Sub Clear()
    If IsArrayAllocated(p_Array) Then Erase p_Array
End Sub

' Helper Methods

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

Private Sub QuickSort( _
    ByRef arrArray() As Variant, _
    ByVal intLoBound As Integer, _
    ByVal intHiBound As Integer _
    )
    
    Dim varPivot As Variant, _
        intLoSwap As Integer, _
        intHiSwap As Integer, _
        varTemp As Variant

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
        While intLoSwap < intHiSwap And arrArray(intLoSwap) <= varPivot
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

Private Sub Class_Terminate()
    Me.Clear
End Sub
