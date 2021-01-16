VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vba_JSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_objScriptHost As Object, _
        p_objScriptEngine As Object

Private Sub Class_Initialize()
    Set p_objScriptHost = CreateObject("HTMLFile")
    Set p_objScriptEngine = p_objScriptHost.parentWindow

    With p_objScriptEngine
            .execScript "var json = {};"
            .execScript "function GetKeys() { var keys = []; for (var k in json) { keys.push(k); } return keys; }", "JScript"
            .execScript "function GetItems() { var items = []; for(var k in json) { items.push(json[k]); } return items; }", "JScript"
            .execScript "function GetItemValue(key) { return json[key]; }", "JScript"
            .execScript "function GetItemType(key) { return GetType(json[key]); }", "JScript"
            .execScript "function GetType(obj) { if (obj === null) { return 'null'; } else if (IsObject(obj)) { return 'object'; } else if (IsArray(obj)) { return 'array'; } else { return typeof(obj); } }", "JScript"
            .execScript "function GetCount() { var count = 0; for (k in json) { if(json.hasOwnProperty(k)) { count++; } } return count; }", "JScript"
            .execScript "function IsObject(obj) { if (obj === null) { return false; } else { return Object.prototype.toString.call(obj) === '[object Object]'; } }", "JScript"
            .execScript "function IsArray(obj) { return Object.prototype.toString.call(obj) === '[object Array]'; }", "JScript"
            .execScript "function ArrayContains(arr, obj) { var i = arr.length; while (i--) { if (arr[i] === obj) { return true; } } return false; }", "JScript"
            .execScript "function GetArrayLength(jsonArr) { return jsonArr.length; }", "JScript"
            .execScript "function GetArrayItem(jsonArr, i) { return jsonArr[i]; }", "JScript"
            .execScript "function GetArrayItemType(jsonArr, i) { return GetType(jsonArr[i]); }", "JScript"
            .execScript "function Exists(key, value, deep) { if (deep) { return SearchContents(json, key, value).length > 0; } else { return json[key] == value; } }", "JScript"
            .execScript "function ValueExists(value, deep) { if (deep) { return SearchContents(json, '', value).length > 0; } else { for (var k in json) { if (json[k] == value) { return true; } } return false; } }", "JScript"
            .execScript "function KeyExists(key, deep) { if (deep) { return SearchContents(json, key, '').length > 0; } else { return key in json; } }", "JScript"
            .execScript "function Find(key, value) { return SearchContents(json, key, value)[0]; }", "JScript"
            .execScript "function FindKey(key) { return SearchContents(json, key, '')[0]; }", "JScript"
            .execScript "function FindValue(value) { return SearchContents(json, '', value)[0]; }", "JScript"
            .execScript "function FindAll(key, value) { return SearchContents(json, key, value); }", "JScript"
            .execScript "function FindAllKeys(key) { return SearchContents(json, key, ''); }", "JScript"
            .execScript "function FindAllValues(value) { return SearchContents(json, '', value); }", "JScript"
            .execScript "function SearchContents(obj, key, value) { var resultSet = []; for (var i in obj) { if (!obj.hasOwnProperty(i)) continue; if (i == key && obj[i] == value || i == key && value == '') { resultSet.push(obj); } else if (obj[i] == value && key == ''){ if (!ArrayContains(resultSet, obj)) { resultSet.push(obj); } } if (typeof(obj[i]) == 'object') { resultSet = resultSet.concat(SearchContents(obj[i], key, value)); } } return resultSet; }", "JScript"
            .execScript "function DeleteItem(key) { delete json[key]; }", "JScript"
            .execScript "function Clear() { for(var member in json) { delete json[member]; } json = {}; }", "JScript"
            .execScript "function Stringify(jsonObj) { var t = typeof(jsonObj); if (t != ""object"" || jsonObj === null) { if (t == ""string"") jsonObj = '""'+jsonObj+'""'; return String(jsonObj); } else { var n, v, jsonStr = [], arr = (jsonObj && jsonObj.constructor == Array); for (n in jsonObj) { v = jsonObj[n]; t = typeof(v); if (t == ""string"") v = '""'+v.replace(/""/g, '\\""').replace(/\r?\n|\r/g, '')+'""'; else if (t == ""object"" && v !== null) v = Stringify(v); jsonStr.push((arr ? """" : '""' + n + '"":') + String(v)); } return (arr ? ""["" : ""{"") + String(jsonStr) + (arr ? ""]"" : ""}""); } };", "JScript"
    End With
End Sub


' Properties


Public Property Get Item( _
    ByVal strKey As String _
    )
    
    If p_objScriptEngine.GetItemType(strKey) = "object" Then
        Set Item = Deserialize(p_objScriptEngine.GetItemValue(strKey))
    Else
        Item = Deserialize(p_objScriptEngine.GetItemValue(strKey))
    End If
End Property

Public Property Get Items()
    Items = Deserialize(p_objScriptEngine.GetItems())
End Property

Public Property Get Keys()
    Keys = Deserialize(p_objScriptEngine.GetKeys())
End Property

Public Property Get Count()
    Count = p_objScriptEngine.GetCount()
End Property


' Methods


Public Sub Clear()
    p_objScriptEngine.Clear
End Sub

Public Function Exists( _
    ByVal strKey As String, _
    ByVal varValue As Variant, _
    ByVal blnDeep As Boolean _
    )
    
    Exists = p_objScriptEngine.Exists(strKey, varValue, blnDeep)
End Function

Public Function KeyExists( _
    ByVal strKey As String, _
    ByVal blnDeep As Boolean _
    )
    
    KeyExists = p_objScriptEngine.KeyExists(strKey, blnDeep)
End Function

Public Function ValueExists( _
    ByVal varValue As Variant, _
    ByVal blnDeep As Boolean _
    )
    
    ValueExists = p_objScriptEngine.ValueExists(varValue, blnDeep)
End Function

Public Function Find( _
    ByVal strKey As String, _
    ByVal varValue As Variant _
    )
    
    Set Find = Deserialize(p_objScriptEngine.Find(strKey, varValue))
End Function

Public Function FindKey( _
    ByVal strKey As String _
    )
    
    Set FindKey = Deserialize(p_objScriptEngine.FindKey(strKey))
End Function

Public Function FindValue( _
    ByVal varValue As Variant _
    )
    
    If p_objScriptEngine.GetType(p_objScriptEngine.FindValue(varValue)) = "object" Then
        Set FindValue = Deserialize(p_objScriptEngine.FindValue(varValue))
    Else
        FindValue = Deserialize(p_objScriptEngine.FindValue(varValue))
    End If
End Function

Public Function FindAll( _
    ByVal strKey As String, _
    ByVal varValue As Variant _
    )
    
    FindAll = Deserialize(p_objScriptEngine.FindAll(strKey, varValue))
End Function

Public Function FindAllKeys( _
    ByVal strKey As String _
    )
    
    FindAllKeys = Deserialize(p_objScriptEngine.FindAllKeys(strKey))
End Function

Public Function FindAllValues( _
    ByVal varValue As Variant _
    )
    
    FindAllValues = Deserialize(p_objScriptEngine.FindAllValues(varValue))
End Function

Public Sub Load( _
    ByVal strFile As String _
    )
    
    Dim objFSO As Object: Set objFSO = CreateObject("Scripting.FileSystemObject")
    FromString objFSO.OpenTextFile(strFile, 1).ReadAll()
    Set FSO = Nothing
End Sub

Public Sub Save( _
    ByVal strFilename As String _
    )
    
    Dim objFSO As Object, _
        objJsonFile As Object

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If objFSO.FileExists(strFilename) Then
        Set objJsonFile = objFSO.OpenTextFile(strFilename, 2, True)
    Else
        Set objJsonFile = objFSO.CreateTextFile(strFilename, True)
    End If

    With objJsonFile
        .WriteLine Me.ToString()
        .Close
    End With

    Set objFSO = Nothing
    Set objJsonFile = Nothing
End Sub

Public Sub FromString( _
    ByVal strJson As String _
    )

    If TypeName(strJson) = "String" Then p_objScriptEngine.execScript "json = " & strJson, "JScript"
End Sub

Public Function ToString()
    ToString = p_objScriptEngine.Stringify(p_objScriptEngine.json)
End Function


' Helper Methods


Private Function Deserialize( _
    ByVal varItem As Variant _
    )
    
    Select Case p_objScriptEngine.GetType(varItem)
        Case "object":
            Set Deserialize = JSONObject(varItem)
        Case "array":
            Deserialize = JSONArray(varItem)
        Case "null":
            Deserialize = Null
        Case "string", "number", "boolean":
            Deserialize = varItem
    End Select
End Function

Private Function JSONObject( _
    ByVal objJSON As Object _
    )
    
    Dim objJsonObj As Object: Set objJsonObj = New vba_JSON
    objJsonObj.FromString p_objScriptEngine.Stringify(objJSON)
    Set JSONObject = objJsonObj
End Function

Private Function JSONArray( _
    ByVal objJSONArr As Object _
    )
    
    Dim arrArray() As Variant, _
        intArrIndex As Integer

    ReDim arrArray(p_objScriptEngine.GetArrayLength(objJSONArr) - 1)

    For intArrIndex = 0 To UBound(arrArray)
        Select Case p_objScriptEngine.GetArrayItemType(objJSONArr, intArrIndex)
            Case "object":
                Set arrArray(intArrIndex) = JSONObject(p_objScriptEngine.GetArrayItem(objJSONArr, intArrIndex))
            Case "array":
                arrArray(intArrIndex) = JSONArray(p_objScriptEngine.GetArrayItem(objJSONArr, intArrIndex))
            Case "null":
                arrArray(intArrIndex) = Null
            Case "string", "number", "boolean":
                arrArray(intArrIndex) = p_objScriptEngine.GetArrayItem(objJSONArr, intArrIndex)
        End Select
    Next

    JSONArray = arrArray
End Function

Private Sub Class_Terminate()
    Set p_objScriptHost = Nothing
    Set p_objScriptEngine = Nothing
End Sub
