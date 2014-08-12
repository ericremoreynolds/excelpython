Attribute VB_Name = "xlpython_v1back"
Option Explicit

Function PyDict(ParamArray elements())
    Dim els As Variant
    els = elements
    Set PyDict = Py.DictFromArray(els)
End Function

Function PyTuple(ParamArray elements())
    Dim els As Variant
    els = elements
    Set PyTuple = Py.TupleFromArray(els)
End Function

Function PyList(ParamArray elements())
    Dim els As Variant
    els = elements
    Set PyList = Py.ListFromArray(els)
End Function

Function PyBuiltin(method As String, Optional args = Empty, Optional kwargs = Empty)
    Set PyBuiltin = Py.Call(Py.BuiltIn, method, args, kwargs)
End Function

Function PyStr(obj) As String
    PyStr = Py.Str(obj)
End Function

Function PyRepr(obj) As String
    PyRepr = PyStr(PyBuiltin("repr", PyTuple(obj)))
End Function

Function PyVar(obj)
    PyVar = Py.Var(obj)
End Function

Function PyObj(value)
    Set PyObj = Py.obj(value)
End Function

Function PyLen(obj) As Long
    PyLen = PyVar(PyBuiltin("len", PyTuple(obj)))
End Function

Function PyType(obj As Variant)
    Set PyType = PyBuiltin("type", PyTuple(obj))
End Function

Function PyCall(instance, Optional method As String = "", Optional args = Empty, Optional kwargs = Empty, Optional console As Boolean = False)
    Set PyCall = Py.Call(instance, method, args, kwargs)
End Function

Function PyGetAttr(instance, attr As String)
    PyGetAttr = Py.GetAttr(instance, attr)
End Function

Sub PySetAttr(instance, attr As String, value)
    Py.SetAttr instance, attr, value
End Sub

Sub PyDelAttr(instance, attr As String)
    Py.DelAttr instance, attr
End Sub

Function PyHasAttr(instance, attr As String) As Boolean
    PyHasAttr = Py.HasAttr(instance, attr)
End Function

Function PyGetItem(instance, key)
    PyGetItem = Py.GetItem(instance, key)
End Function

Sub PySetItem(instance, key, value)
   Py.SetItem instance, key, value
End Sub

Sub PyDelItem(instance, key)
    Py.DelItem instance, key
End Sub

Function PyContains(instance, key) As Boolean
    PyContains = Py.Contains(instance, key)
End Function

Function PyEval(expression As String, Optional locals = Empty, Optional globals = Empty, Optional addpath As String = "", Optional path As String = "")
    Set PyEval = Py.Eval(expression, locals, globals)
End Function

Function PyExec(statement As String, Optional locals = Empty, Optional globals = Empty, Optional addpath As String = "", Optional path As String = "")
    Set PyExec = Py.Exec(statement, locals, globals)
End Function

Function PyModule(name As String, Optional reload As Boolean = True, Optional addpath As String = "", Optional path As String = "")
    Set PyModule = Py.Module(name)
End Function
