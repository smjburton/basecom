Option Explicit

' Static functions for helping with URI operations.

' RelativeTo() - compares two paths and makes one relative to the other
' AbsoluteTo() - makes a relative path absolute based on another path
' Equals() - determines if the given URLs are the same - disregarding default ports, capitalization, dot-pathnames, query-parameter order, etc.
' WithinString(strURI) - Detects URIs within random text
' CommonPath() - determines the common base directory of two paths.
' Parse(string url) - parses a string into its URI components. returns an object containing the found components
' ParseAuthority(string url, object parts) - parses a string's beginning into its URI components username, password, hostname, port. Found components are appended to the parts parameter. Remaining string is returned
' ParseUserinfo(string url, object parts) - parses a string's beginning into its URI components username, password. Found components are appended to the parts parameter. Remaining string is returned
' ParseHost(string url, object parts) - parses a string's beginning into its URI components hostname, port. Found components are appended to the parts parameter. Remaining string is returned
' ParseQuery(string querystring) - parses the passed query string into an object. Returns object {propertyName: propertyValue}
' Build(object parts) - serializes the URI components passed in parts into a URI string
' BuildAuthority(object parts) - serializes the URI components username, password, hostname, port passed in parts into a URI string
' BuildUserinfo(object parts) - serializes the URI components username, password passed in parts into a URI string
' BuildHost(object parts) - serializes the URI components hostname, port passed in parts into a URI string
' BuildQuery(object data, [boolean duplicateQueryParameters], [boolean escapeQuerySpace]) - serializes the query string parameters

If WScript.ScriptName = "v_Util_URI.vbs" Then

End If