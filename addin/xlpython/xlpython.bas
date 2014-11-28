Attribute VB_Name = "xlpython"
Option Private Module
Option Explicit

#If VBA7 Then
        #If win64 Then
                Const XLPyDLLName As String = "xlpython64-b1.dll"
                Private Declare PtrSafe Function XLPyDLLActivate Lib "xlpython64-b1.dll" (ByRef result As Variant, ByVal config As String, ByVal activation As Long) As Long
                Private Declare PtrSafe Function XLPyDLLNDims Lib "xlpython64-b1.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
                Private Declare PtrSafe Function XLPyDLLVersion Lib "xlpython64-b1.dll" (ByRef tag as String, ByRef version as Double, ByRef architecture as String) As Long
        #Else
                Private Const XLPyDLLName As String = "xlpython32-b1.dll"
                Private Declare PtrSafe Function XLPyDLLActivate Lib "xlpython32-b1.dll" (ByRef result As Variant, ByVal config As String, ByVal activation As Long) As Long
                Private Declare PtrSafe Function XLPyDLLNDims Lib "xlpython32-b1.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
                Private Declare PtrSafe Function XLPyDLLVersion Lib "xlpython32-b1.dll" (ByRef tag as String, ByRef version as Double, ByRef architecture as String) As Long
        #End If
        Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
#Else
        Private Const XLPyDLLName As String = "xlpython32-b1.dll"
        Private Declare Function XLPyDLLActivate Lib "xlpython32-b1.dll" (ByRef result As Variant, ByVal config As String, ByVal activation As Long) As Long
        Private Declare Function XLPyDLLNDims Lib "xlpython32-b1.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
        Private Declare Function XLPyDLLVersion Lib "xlpython32-b1.dll" (ByRef tag As String, ByRef version As Double, ByRef architecture As String) As Long
        Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
#End If

Private Function XLPyVersion() As String
    XLPyVersion = "2.1.0"
End Function

Function XLPyConfig() As String
    XLPyConfig = ThisWorkbook.Path + "\xlpython\xlpython.cfg"
End Function

Sub XLPyLoadDLL()
    LoadLibrary ThisWorkbook.Path + "\xlpython\" + XLPyDLLName
End Sub

Function NDims(ByRef src As Variant, dims As Long, Optional transpose As Boolean = False)
    XLPyLoadDLL
    If 0 <> XLPyDLLNDims(src, dims, transpose, NDims) Then Err.Raise 1001, Description:=NDims
End Function

Function Py(Optional activation As Long = 1)
    XLPyLoadDLL
        If 0 <> XLPyDLLActivate(Py, XLPyConfig, activation) Then Err.Raise 1000, Description:=Py
End Function
