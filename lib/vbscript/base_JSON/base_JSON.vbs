Option Explicit

' Add
' Remove
' Splice

Include "base_Sys_Script"

Class base_JSON
	Private p_objScript

	Private Sub Class_Initialize()
		Set p_objScript = New base_Sys_Script

		With p_objScript
			.Language = "JScript"
			.AddCode("var json = {};")
			.AddCode("function getKeys() { var keys = []; for (var k in json) { keys.push(k); } return keys; }")
			.AddCode("function getItems() { var items = []; for(var k in json) { items.push(json[k]); } return items; }")
			.AddCode("function getItemValue(key) { return json[key]; }")
			.AddCode("function getItemType(key) { return getType(json[key]); }")
			.AddCode("function getType(obj) { if (obj === null) { return 'null'; } else if (isObject(obj)) { return 'object'; } else if (isArray(obj)) { return 'array'; } else { return typeof(obj); } }")
			.AddCode("function getCount() { var count = 0; for (k in json) { if(json.hasOwnProperty(k)) { count++; } } return count; }")
			.AddCode("function isObject(obj) { if (obj === null) { return false; } else { return Object.prototype.toString.call(obj) === '[object Object]'; } }")
			.AddCode("function isArray(obj) { return Object.prototype.toString.call(obj) === '[object Array]'; }")
			.AddCode("function arrayContains(arr, obj) { var i = arr.length; while (i--) { if (arr[i] === obj) { return true; } } return false; }")
			.AddCode("function getArrayLength(jsonArr) { return jsonArr.length; }")
			.AddCode("function getArrayItem(jsonArr, i) { return jsonArr[i]; }")
			.AddCode("function getArrayItemType(jsonArr, i) { return getType(jsonArr[i]); }")
			.AddCode("function exists(key, value, deep) { if (deep) { return searchContents(json, key, value).length > 0; } else { return json[key] == value; } }")
			.AddCode("function valueExists(value, deep) { if (deep) { return searchContents(json, '', value).length > 0; } else { for (var k in json) { if (json[k] == value) { return true; } } return false; } }")
			.AddCode("function keyExists(key, deep) { if (deep) { return searchContents(json, key, '').length > 0; } else { return key in json; } }")
			.AddCode("function find(key, value) { return searchContents(json, key, value)[0]; }")
			.AddCode("function findKey(key) { return searchContents(json, key, '')[0]; }")
			.AddCode("function findValue(value) { return searchContents(json, '', value)[0]; }")
			.AddCode("function findAll(key, value) { return searchContents(json, key, value); }")
			.AddCode("function findAllKeys(key) { return searchContents(json, key, ''); }")
			.AddCode("function findAllValues(value) { return searchContents(json, '', value); }")
			.AddCode("function searchContents(obj, key, value) { var resultSet = []; for (var i in obj) { if (!obj.hasOwnProperty(i)) continue; if (i == key && obj[i] == value || i == key && value == '') { resultSet.push(obj); } else if (obj[i] == value && key == ''){ if (!arrayContains(resultSet, obj)) { resultSet.push(obj); } } if (typeof(obj[i]) == 'object') { resultSet = resultSet.concat(searchContents(obj[i], key, value)); } } return resultSet; }")
			.AddCode("function deleteItem(key) { delete json[key]; }")
			.AddCode("function clear() { for(var member in json) { delete json[member]; } json = {}; }")
			.AddCode("function stringify(jsonObj) { var t = typeof(jsonObj); if (t != ""object"" || jsonObj === null) { if (t == ""string"") jsonObj = '""'+jsonObj+'""'; return String(jsonObj); } else { var n, v, jsonStr = [], arr = (jsonObj && jsonObj.constructor == Array); for (n in jsonObj) { v = jsonObj[n]; t = typeof(v); if (t == ""string"") v = '""'+v.replace(/""/g, '\\""').replace(/\r?\n|\r/g, '')+'""'; else if (t == ""object"" && v !== null) v = stringify(v); jsonStr.push((arr ? """" : '""' + n + '"":') + String(v)); } return (arr ? ""["" : ""{"") + String(jsonStr) + (arr ? ""]"" : ""}""); } };")
		End With
	End Sub


	' Properties


	Public Default Property Get Item(strKey)
		If p_objScript.Run("getItemType", Array(strKey)) = "object" Then
			Set Item = Deserialize(p_objScript.Run("getItemValue", Array(strKey)))
		Else
			Item = Deserialize(p_objScript.Run("getItemValue", Array(strKey)))
		End If
	End Property

	Public Property Get Items()
		Items = Deserialize(p_objScript.Run("getItems", Array()))
	End Property

	Public Property Get Keys()
		Keys = Deserialize(p_objScript.Run("getKeys", Array()))
	End Property

	Public Property Get Count()
		Count = p_objScript.Run("getCount", Array())
	End Property

	
	' Methods


	Public Sub Clear()
		p_objScript.Run "clear", Array()
	End Sub

	Public Function Exists(strKey, varValue, blnDeep)
		Exists = p_objScript.Run("exists", Array(strKey, varValue, blnDeep))
	End Function

	Public Function KeyExists(strKey, blnDeep)
		KeyExists = p_objScript.Run("keyExists", Array(strKey, blnDeep))
	End Function

	Public Function ValueExists(varValue, blnDeep)
		ValueExists = p_objScript.Run("valueExists", Array(varValue, blnDeep))
	End Function

	Public Function Find(strKey, varValue)
		Set Find = Deserialize(p_objScript.Run("find", Array(strKey, varValue)))
	End Function

	Public Function FindKey(strKey)
		Set FindKey = Deserialize(p_objScript.Run("findKey", Array(strKey)))
	End Function

	Public Function FindValue(varValue)
		If p_objScript.Run("getType", Array(p_objScript.Run("findValue", Array(varValue)))) = "object" Then
			Set FindValue = Deserialize(p_objScript.Run("findValue", Array(varValue)))
		Else
			FindValue = Deserialize(p_objScript.Run("findValue", Array(varValue)))
		End If
	End Function

	Public Function FindAll(strKey, varValue)
		FindAll = Deserialize(p_objScript.Run("findAll", Array(strKey, varValue)))
	End Function

	Public Function FindAllKeys(strKey)
		FindAllKeys = Deserialize(p_objScript.Run("findAllKeys", Array(strKey)))
	End Function

	Public Function FindAllValues(varValue)
		FindAllValues = Deserialize(p_objScript.Run("findAllValues", Array(varValue)))
	End Function

	Public Sub Load(strFile)
		Dim objFSO
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Me.FromString objFSO.OpenTextFile(strFile, 1).ReadAll()
		Set FSO = Nothing
	End Sub

	Public Sub Save(strFilename)
		Dim objFSO, _
			objJsonFile

		Set objFSO = CreateObject("Scripting.FileSystemObject")

		If objFSO.FileExists(strFilename) Then
			Set objJsonFile = objFSO.OpenTextFile(strFilename, 2, True)			
		Else
			Set objJsonFile = objFSO.CreateTextFile(strFilename, True)
		End If

		With objJsonFile
			.WriteLine Me.ToString()
			.Close()
		End With

		Set objFSO = Nothing
		Set objJsonFile = Nothing
	End Sub

	Public Sub FromString(strJSON)
		If TypeName(strJSON) = "String" Then p_objScript.Variable("json") = strJSON
	End Sub

	Public Function ToString()
		ToString = p_objScript.Run("stringify", Array(p_objScript.Variable("json")))
	End Function

	
	' Helper Methods


	Private Function Deserialize(varItem)
		Select Case p_objScript.Run("getType", Array(varItem))
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

	Private Function JSONObject(objJSON)
		Dim objJsonObj
		Set objJsonObj = New vbs_JSON
		objJsonObj.FromString p_objScript.Run("stringify", Array(objJSON))
		Set JSONObject = objJsonObj
	End Function

	Private Function JSONArray(objJSONArr)
		Dim arrArray(), _
			i

		ReDim arrArray(p_objScript.Run("getArrayLength", Array(objJSONArr)) - 1)

		For i = 0 To UBound(arrArray)
			Select Case p_objScript.Run("getArrayItemType", Array(objJSONArr, i))
				Case "object":
					Set arrArray(i) = JSONObject(p_objScript.Run("getArrayItem", Array(objJSONArr, i)))
				Case "array":
					arrArray(i) = JSONArray(p_objScript.Run("getArrayItem", Array(objJSONArr, i)))
				Case "null":
					arrArray(i) = Null
				Case "string", "number", "boolean":
					arrArray(i) = p_objScript.Run("getArrayItem", Array(objJSONArr, i))
			End Select
		Next

		JSONArray = arrArray
	End Function

	Private Sub Class_Terminate()
		Set p_objScript = Nothing
	End Sub 
End Class

If WScript.ScriptName = "base_JSON.vbs" Then
	Dim json
	Set json = New base_JSON

	json.FromString "{""key1"": null, ""key2"": { ""key3"": ""val3"" }, " & _
			"""key4"": ""val4"", ""key5"": true, ""key6"": 7.8, " & _
	 		"""employees"":[ { ""firstName"":""John"", ""lastName""" & _
	 		":""Doe"" }, { ""firstName"":""Anna"", ""lastName"":" & _
	 		"""Smith"" }, { ""firstName"":""Peter"", ""lastName"":" & _
	 		"""Jones"" } ] }"

	WScript.Echo UBound(json.FindAllValues("val4"))

	' json.FromString "{""array"": [ ""val1"", 2, true, null, { ""firstName"":""Bob"" }, [ ""val1"", ""val2"", ""val3"", { ""key1"":""val2"" }, [ [ [ { ""aTestKey"" : ""aTestVal"" } ], { ""firstName"" : ""Joe"" } ] ] ] ] }"
End If
