Attribute VB_Name = "ExcelPython"
#If VBA7 Then

Private Declare PtrSafe Function GetTempPath32 Lib "kernel32" _
   Alias "GetTempPathA" (ByVal nBufferLength As LongPtr, _
   ByVal lpBuffer As String) As Long
   
Private Declare PtrSafe Function GetTempFileName32 Lib "kernel32" _
   Alias "GetTempFileNameA" (ByVal lpszPath As String, _
   ByVal lpPrefixString As String, ByVal wUnique As Long, _
   ByVal lpTempFileName As String) As Long

#Else

Private Declare Function GetTempPath32 Lib "kernel32" _
   Alias "GetTempPathA" (ByVal nBufferLength As Long, _
   ByVal lpBuffer As String) As Long
   
Private Declare Function GetTempFileName32 Lib "kernel32" _
   Alias "GetTempFileNameA" (ByVal lpszPath As String, _
   ByVal lpPrefixString As String, ByVal wUnique As Long, _
   ByVal lpTempFileName As String) As Long

#End If
   
Private Function GetTempFileName()
    Dim sTmpPath As String * 512
    Dim sTmpName As String * 576
    Dim nRet As Long
    nRet = GetTempPath32(512, sTmpPath)
    If nRet = 0 Then Err.Raise 1234, Description:="GetTempPath failed."
    nRet = GetTempFileName32(sTmpPath, "vba", 0, sTmpName)
    If nRet = 0 Then Err.Raise 1234, Description:="GetTempFileName failed."
    GetTempFileName = Left$(sTmpName, InStr(sTmpName, vbNullChar) - 1)
End Function

Function ModuleIsPresent(ByVal wb As Workbook, moduleName As String) As Boolean
    On Error GoTo not_present
    Set x = wb.VBProject.VBComponents.Item(moduleName)
    ModuleIsPresent = True
    Exit Function
not_present:
    ModuleIsPresent = False
End Function


Sub SetupExcelPython(control As IRibbonControl)
    Set wb = ActiveWorkbook
    If wb.Path = "" Then
        MsgBox "Please save this workbook first, as a macro-enabled workbook."
        Exit Sub
    End If
    If LCase$(Right$(wb.name, 5)) <> ".xlsm" And LCase$(Right$(wb.name, 5)) <> ".xlsb" Then
        MsgBox "Please save this workbook, " + wb.name + ", as a macro-enabled workbook first."
        Exit Sub
    End If
    mssg = "This action will:" _
        + vbCrLf + "  - install the ExcelPython runtime in the folder '" + wb.Path + Application.PathSeparator + "xlpython'" _
        + vbCrLf + "  - set up this workbook ('" + wb.name + "') to interact with Python" _
        + vbCrLf + vbCrLf + "Do you want to proceed?"
    If vbYes = MsgBox(mssg, vbYesNo, "Set up workbook for ExcelPython") Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(wb.Path + Application.PathSeparator + "xlpython") Then
            isVersionOK = False
            ver = "?.?.?"
            For Each f In fso.GetFolder(ThisWorkbook.Path + Application.PathSeparator + "xlpython").Files
                If LCase$(Right$(f, 4)) = ".dll" Then
                    isVersionOK = fso.FileExists(wb.Path + Application.PathSeparator + "xlpython" + Application.PathSeparator + fso.GetFileName(f))
                    ver = Mid$(fso.GetBaseName(f), InStr(fso.GetBaseName(f), "-") + 1)
                    Exit For
                End If
            Next f
            If Not isVersionOK Then
                MsgBox "The installation folder already exists, but it does not contain ExcelPython version " + ver + "." _
                    + vbCrLf + vbCrLf + "Installation folder: " + wb.Path + Application.PathSeparator + "xlpython" _
                    + vbCrLf + vbCrLf + "To set up a fresh install please delete it and try again. Note that you may need to close Excel to delete it." _
                    , vbCritical, "Error installing ExcelPython runtime"
                Exit Sub
            End If
        Else
            fso.CopyFolder ThisWorkbook.Path + Application.PathSeparator + "xlpython", wb.Path + Application.PathSeparator + "xlpython"
        End If
        
        On Error GoTo not_present
        wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents("xlpython")
not_present:
        On Error GoTo 0
        wb.VBProject.VBComponents.Import wb.Path + Application.PathSeparator + "xlpython" + Application.PathSeparator + "xlpython.bas"
        
        ' create skeleton py file
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FileExists(wb.Path + Application.PathSeparator + fso.GetBaseName(wb.name) + ".py") Then
            Set f = fso.CreateTextFile(wb.Path + Application.PathSeparator + fso.GetBaseName(wb.name) + ".py", True)
            f.WriteLine "from xlpython import *"
            f.Close
        End If
        'MsgBox "You can now write user-defined functions for this workbook in Python in the file '" + wb.Path + Application.PathSeparator + fso.GetBaseName(wb.Name) + ".py'." + vbCrLf + "Please consult the online docs for more information on how it works.", Title:="ExcelPython setup successful!"
    End If
End Sub

Sub XLPMacroOptions2010(macroName As String, desc, argdescs() As String)
    Application.MacroOptions macroName, Description:=desc, ArgumentDescriptions:=argdescs
End Sub

Sub ImportPythonUDFs(control As IRibbonControl)
    sTab = "    "

    Set wb = ActiveWorkbook
    If Not ModuleIsPresent(wb, "xlpython") Then
        MsgBox "The active workbook does not seem to have been set up to use ExcelPython yet."
        Exit Sub
    End If
    
    Set Py = Application.Run("'" + wb.name + "'!Py")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filename = GetTempFileName()
    Set f = fso.CreateTextFile(filename, True)
    f.WriteLine "Attribute VB_Name = ""xlpython_udfs"""
    f.WriteLine
    f.WriteLine "Function PyScriptPath() As String"
    f.WriteLine sTab + "PyScriptPath = Left$(ThisWorkbook.Name, Len(ThisWorkbook.Name)-5) ' assume that it ends in .xlsm"
    ' f.WriteLine sTab + "If LCase$(Right$(PyScriptPath, 5)) = "".xlsm"" Then PyScriptPath = Left$(PyScriptPath, Len(PyScriptPath)-5)"
    f.WriteLine sTab + "PyScriptPath = ThisWorkbook.Path + Application.PathSeparator + PyScriptPath + "".py"""
    f.WriteLine "End Function"
    f.WriteLine
    
    Dim scriptPath As String
    scriptPath = wb.Path + Application.PathSeparator + fso.GetBaseName(wb.name) + ".py"
    Set scriptVars = Py.Call(Py.Module("xlpython"), "udf_script", Py.Tuple(scriptPath))
    For Each svar In Py.Call(scriptVars, "values")
        If Py.HasAttr(svar, "__xlfunc__") Then
            Set xlfunc = Py.GetAttr(svar, "__xlfunc__")
            Set xlret = Py.GetItem(xlfunc, "ret")
            fname = Py.Str(Py.GetItem(xlfunc, "name"))
            
            Dim ftype As String
            If Py.Var(Py.GetItem(xlfunc, "sub")) Then ftype = "Sub" Else ftype = "Function"
            
            f.Write ftype + " " + fname + "("
            first = True
            vararg = ""
            nArgs = Py.Len(Py.GetItem(xlfunc, "args"))
            For Each arg In Py.GetItem(xlfunc, "args")
                If Not Py.Bool(Py.GetItem(arg, "vba")) Then
                    argname = Py.Str(Py.GetItem(arg, "name"))
                    If Not first Then f.Write ", "
                    If Py.Bool(Py.GetItem(arg, "vararg")) Then
                        f.Write "ParamArray "
                        vararg = argname
                    End If
                    f.Write argname
                    If Py.Bool(Py.GetItem(arg, "vararg")) Then
                        f.Write "()"
                    End If
                    first = False
                End If
            Next arg
            f.WriteLine ")"
            If ftype = "Function" Then
                f.WriteLine sTab + "If TypeOf Application.Caller Is Range Then On Error GoTo failed"
            End If
            
            If vararg <> "" Then
                f.WriteLine sTab + "ReDim argsArray(1 to UBound(" + vararg + ") - LBound(" + vararg + ") + " + CStr(nArgs) + ")"
            End If
            j = 1
            For Each arg In Py.GetItem(xlfunc, "args")
                If Not Py.Bool(Py.GetItem(arg, "vba")) Then
                    argname = Py.Str(Py.GetItem(arg, "name"))
                    If Py.Bool(Py.GetItem(arg, "vararg")) Then
                        f.WriteLine sTab + "For k = lbound(" + vararg + ") to ubound(" + vararg + ")"
                        argname = vararg + "(k)"
                    End If
                    If Not Py.Var(Py.GetItem(arg, "range")) Then
                        f.WriteLine sTab + "If TypeOf " + argname + " Is Range Then " + argname + " = " + argname + ".Value2"
                    End If
                    dims = Py.Var(Py.GetItem(arg, "dims"))
                    marshal = Py.Str(Py.GetItem(arg, "marshal"))
                    If dims <> -1 Or marshal = "nparray" Or marshal = "list" Then
                        f.WriteLine sTab + "If Not TypeOf " + argname + " Is Object Then"
                        If dims <> -1 Then
                            f.WriteLine sTab + sTab + argname + " = NDims(" + argname + ", " + CStr(dims) + ")"
                        End If
                        If marshal = "nparray" Then
                            dtype = Py.Var(Py.GetItem(arg, "dtype"))
                            If IsNull(dtype) Then
                                f.WriteLine sTab + sTab + "Set " + argname + " = Py.Call(Py.Module(""numpy""), ""array"", Py.Tuple(" + argname + "))"
                            Else
                                f.WriteLine sTab + sTab + "Set " + argname + " = Py.Call(Py.Module(""numpy""), ""array"", Py.Tuple(" + argname + ", """ + dtype + """))"
                            End If
                        ElseIf marshal = "list" Then
                            f.WriteLine sTab + sTab + "Set " + argname + " = Py.Call(Py.Eval(""lambda t: [ list(x) if isinstance(x, tuple) else x for x in t ] if isinstance(t, tuple) else t""), Py.Tuple(" + argname + "))"
                        End If
                        f.WriteLine sTab + "End If"
                    End If
                    If Py.Bool(Py.GetItem(arg, "vararg")) Then
                        f.WriteLine sTab + "argsArray(" + CStr(j) + " + k - LBound(" + vararg + ")) = " + argname
                        f.WriteLine sTab + "Next k"
                    Else
                        If vararg <> "" Then
                            f.WriteLine sTab + "argsArray(" + CStr(j) + ") = " + argname
                            j = j + 1
                        End If
                    End If
                End If
            Next arg
            
            If vararg <> "" Then
                f.WriteLine sTab + "Set args = Py.TupleFromArray(argsArray)"
            Else
                f.Write sTab + "Set args = Py.Tuple("
                first = True
                For Each arg In Py.GetItem(xlfunc, "args")
                    If Not first Then f.Write ", "
                    If Not Py.Bool(Py.GetItem(arg, "vba")) Then
                        f.Write Py.Str(Py.GetItem(arg, "name"))
                    Else
                        f.Write Py.Str(Py.GetItem(arg, "vba"))
                    End If
                    first = False
                Next arg
                f.WriteLine ")"
            End If
            
            If Py.Bool(Py.GetItem(xlfunc, "xlwings")) Then
                f.WriteLine sTab + "Py.SetAttr Py.GetAttr(Py.Module(""xlwings""), ""xlplatform""), ""xl_app_latest"", Application"
                f.WriteLine sTab + "Py.SetAttr Py.Module(""xlwings.main""), ""xl_workbook_latest"", ThisWorkbook"
            End If
            
            f.WriteLine sTab + "Set xlpy = Py.Module(""xlpython"")"
            f.WriteLine sTab + "Set script = Py.Call(xlpy, ""udf_script"", Py.Tuple(PyScriptPath))"
            f.WriteLine sTab + "Set func = Py.GetItem(script, """ + fname + """)"
            If ftype = "Sub" Then
                f.WriteLine sTab + "Py.Call func, args"
            Else
                f.WriteLine sTab + "Set " + fname + " = Py.Call(func, args)"
                marshal = Py.Str(Py.GetItem(xlret, "marshal"))
                Select Case marshal
                Case "auto"
                    f.WriteLine sTab + "If TypeOf Application.Caller Is Range Then " + fname + " = Py.Var(" + fname + ", " + Py.Str(Py.GetItem(xlret, "lax")) + ")"
                Case "var"
                    f.WriteLine sTab + fname + " = Py.Var(" + fname + ", " + Py.Str(Py.GetItem(xlret, "lax")) + ")"
                Case "str"
                    f.WriteLine sTab + fname + " = Py.Str(" + fname + ")"
                End Select
            End If
            
            If ftype = "Function" Then
                f.WriteLine sTab + "Exit " + ftype
                f.WriteLine "failed:"
                f.WriteLine sTab + fname + " = Err.Description"
            End If
            f.WriteLine "End " + ftype
            f.WriteLine
        End If
    Next svar
    f.Close
    
    On Error GoTo not_present
    wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents("xlpython_udfs")
not_present:
    On Error GoTo 0
    wb.VBProject.VBComponents.Import filename
    
    For Each svar In Py.Call(scriptVars, "values")
        If Py.HasAttr(svar, "__xlfunc__") Then
            Set xlfunc = Py.GetAttr(svar, "__xlfunc__")
            Set xlret = Py.GetItem(xlfunc, "ret")
            Set xlargs = Py.GetItem(xlfunc, "args")
            fname = Py.Str(Py.GetItem(xlfunc, "name"))
            fdoc = Py.Str(Py.GetItem(xlret, "doc"))
            nArgs = 0
            For Each arg In xlargs
                If Not Py.Bool(Py.GetItem(arg, "vba")) Then nArgs = nArgs + 1
            Next arg
            If nArgs > 0 And Val(Application.Version) >= 14 Then
                ReDim argdocs(1 To WorksheetFunction.Max(1, nArgs)) As String
                nArgs = 0
                For Each arg In xlargs
                    If Not Py.Bool(Py.GetItem(arg, "vba")) Then
                        nArgs = nArgs + 1
                        argdocs(nArgs) = Left$(Py.Str(Py.GetItem(arg, "doc")), 255)
                    End If
                Next arg
                XLPMacroOptions2010 "'" + wb.name + "'!" + fname, Left$(fdoc, 255), argdocs
            Else
                Application.MacroOptions "'" + wb.name + "'!" + fname, Description:=Left$(fdoc, 255)
            End If
        End If
    Next svar
    
    'MsgBox "Import successful!"
End Sub



