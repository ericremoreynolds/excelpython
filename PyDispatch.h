#include <Python.h>
#include <Windows.h>

PyObject* PyDispatch_New(IDispatch* pDispatch);

PyObject* PyProgressCallback_New(IProgressCallback* pCallback);

_COM_SMARTPTR_TYPEDEF(IProgressCallback, __uuidof(IProgressCallback));