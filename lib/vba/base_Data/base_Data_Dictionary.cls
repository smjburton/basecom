VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "base_Data_Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_Dictionary As Object

Private Sub Class_Initialize()
   Set p_Dictionary = CreateObject("Scripting.Dictionary")
End Sub


' Properties


Public Property Get CompareMode() As Integer
   CompareMode = p_Dictionary.CompareMode
End Property

Public Property Let CompareMode( _
   ByVal intCompareMode As Integer _
   )
   
   If p_Dictionary.Count = 0 Then p_Dictionary.CompareMode = intCompareMode
End Property

Public Property Get Count() As Integer
    Count = p_Dictionary.Count
End Property

' *** Make default property
Public Property Get Item( _
   ByVal strKey As String _
   )
   
   If IsObject(p_Dictionary.Item(strKey)) Then
       Set Item = p_Dictionary.Item(strKey)
   Else
       Item = p_Dictionary.Item(strKey)
   End If
End Property

Public Property Let Item( _
   ByVal strKey As String, _
   ByVal varItem As Variant _
   )
   
   p_Dictionary.Item(strKey) = varItem
End Property

Public Property Set Item( _
   ByVal strKey As String, _
   ByVal varItem As Variant _
   )
   
   Set p_Dictionary.Item(strKey) = varItem
End Property

Public Property Get Items() As Variant()
    Items = p_Dictionary.Items()
End Property

Public Property Let Key( _
    ByVal strKey As String, _
    ByVal strNewKey As String _
    )
    
    p_Dictionary.Key(strKey) = strNewKey
End Property

Public Property Get Keys() As Variant()
    Keys = p_Dictionary.Keys()
End Property


' Methods


Public Sub Add( _
    ByVal strKey As String, _
    ByVal varItem As Variant _
    )
    
    p_Dictionary.Add strKey, varItem
End Sub

Public Function Exists( _
    ByVal strKey As String _
    ) As Boolean
    
    Exists = p_Dictionary.Exists(strKey)
End Function

Public Sub Remove( _
    ByVal strKey As String _
    )
    
    p_Dictionary.Remove strKey
End Sub

Public Sub RemoveAll()
    p_Dictionary.RemoveAll
End Sub

Private Sub Class_Terminate()
    Set p_Dictionary = Nothing
End Sub
