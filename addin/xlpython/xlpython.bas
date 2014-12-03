Attribute VB_Name = "xlpython"
Option Private Module
Option Explicit

#If VBA7 Then
	#If win64 Then
		Const XLPyDLLName As String = "xlpython64-2.0.8.dll"
		Declare PtrSafe Function XLPyDLLActivate Lib "xlpython64-2.0.8.dll" (ByRef result As Variant, Optional ByVal config As String = "") As Long
		Declare PtrSafe Function XLPyDLLNDims Lib "xlpython64-2.0.8.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
	#Else
		Private Const XLPyDLLName As String = "xlpython32-2.0.8.dll"
		Private Declare PtrSafe Function XLPyDLLActivate Lib "xlpython32-2.0.8.dll" (ByRef result As Variant, Optional ByVal config As String = "") As Long
		Private Declare PtrSafe Function XLPyDLLNDims Lib "xlpython32-2.0.8.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
	#End If
	Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
#Else
	#If win64 Then
		Const XLPyDLLName As String = "xlpython64-2.0.8.dll"
		Declare Function XLPyDLLActivate Lib "xlpython64-2.0.8.dll" (ByRef result As Variant, Optional ByVal config As String = "") As Long
		Declare Function XLPyDLLNDims Lib "xlpython64-2.0.8.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
	#Else
		Private Const XLPyDLLName As String = "xlpython32-2.0.8.dll"
		Private Declare Function XLPyDLLActivate Lib "xlpython32-2.0.8.dll" (ByRef result As Variant, Optional ByVal config As String = "") As Long
		Private Declare Function XLPyDLLNDims Lib "xlpython32-2.0.8.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
	#End If
	Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
#EndIf

Private Function XLPyFolder() As String
    XLPyFolder = ThisWorkbook.Path + "\xlpython"
End Function

Function XLPyConfig() As String
    XLPyConfig = XLPyFolder + "\xlpython.cfg"
End Function

Sub XLPyLoadDLL()
    LoadLibrary XLPyFolder + "\" + XLPyDLLName
End Sub

Function NDims(ByRef src As Variant, dims As Long, Optional transpose As Boolean = False)
    XLPyLoadDLL
    If 0 <> XLPyDLLNDims(src, dims, transpose, NDims) Then Err.Raise 1001, Description:=NDims
End Function

Function Py()
    XLPyLoadDLL
    If 0 <> XLPyDLLActivate(Py, XLPyConfig) Then Err.Raise 1000, Description:=Py
End Function