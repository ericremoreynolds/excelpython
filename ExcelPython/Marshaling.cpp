#include "ExcelPython.h"

#undef GetFreeSpace
#import "C:/Windows/System32/scrrun.dll"

void StrToVariant(const std::string& str, VARIANT* var)
{
	LogCtx log("StrToVariant");

	VariantClear(var);

	int sz = (int) str.length() + 1;
	OLECHAR* wide = new OLECHAR[sz];
	MultiByteToWideChar(CP_ACP, 0, str.c_str(), sz * sizeof(OLECHAR), wide, sz);
	var->vt = VT_BSTR;
	var->bstrVal = SysAllocString(wide);
	delete[] wide;
}

std::string WCharToStdString(const wchar_t* str)
{
	LogCtx log("WCharToStdString");

	int len = (int) wcslen(str);
	boost::scoped_array<char> narrow(new char[len+1]);
	WideCharToMultiByte(CP_ACP, 0, str, len, narrow.get(), len+1, "?", NULL);
	narrow[len] = 0;
	std::string ret = narrow.get();
	return ret;
}

std::string BStrToStdString(BSTR str)
{
	LogCtx log("BStrToStdString");

	BOOL bUsedDefaultChar;
	int len = (int) SysStringLen(str);
	boost::scoped_array<char> narrow(new char[len+1]);
	WideCharToMultiByte(CP_ACP, 0, str, len, narrow.get(), len+1, "?", &bUsedDefaultChar);
	narrow[len] = 0;
	std::string ret = narrow.get();
	return ret;
}

BSTR CStrToBStr(const char* str)
{
	LogCtx log("CStrToBStr");

	int sz = (int) strlen(str) + 1;
	OLECHAR* wide = new OLECHAR[sz];
	MultiByteToWideChar(CP_ACP, 0, str, sz * sizeof(OLECHAR), wide, sz);
	BSTR ret = SysAllocString(wide);
	delete[] wide;
	return ret;
}


PyObject* BStrToPyString(BSTR str)
{
	LogCtx log("BStrToPyString");

	BOOL bUsedDefaultChar;
	int len = (int) SysStringLen(str);
	char* narrow = new char[len+1];
	WideCharToMultiByte(CP_ACP, 0, str, len, narrow, len+1, "?", &bUsedDefaultChar);
	narrow[len] = 0;
	PyObject* ret = PyString_FromStringAndSize(narrow, (Py_ssize_t) len);
	delete[] narrow;
	if(ret == NULL)
		throw PythonException();
	return ret;
}

PyObject* IUnknownToPyWin32(IUnknown* pUnk)
{
	LogCtx log("IUnknownToPyWin32");

	PyNewRef pythoncom(PyImport_ImportModule("pythoncom"));				// import pythoncom
	PyNewRef w32client(PyImport_ImportModule("win32com.client"));		// import win32com.client
#ifdef _WIN64
	PyNewRef address(PyLong_FromLongLong((LONG_PTR) pUnk));
#else
	PyNewRef address(PyInt_FromLong((LONG_PTR) pUnk));
#endif
	PyNewRef iunk(PyObject_CallMethod(pythoncom, "ObjectFromAddress", "O", address));
	PyNewRef iid_idispatch(PyObject_GetAttrString(pythoncom, "IID_IDispatch"));
	PyRef idisp;
	try
	{
		idisp = PyNewRef(PyObject_CallMethod(iunk, "QueryInterface", "O", iid_idispatch));
	}
	catch(PythonException&)
	{
		return iunk.detach();
	}

	PyNewRef disp(PyObject_CallMethod(w32client, "Dispatch", "O", idisp));
	return disp.detach();
}

PyObject* VariantToPy(VARIANT* var, int dims)
{
	LogCtx log("VariantToPy");

	switch(var->vt)
	{
	case VT_EMPTY:
		{
			Py_INCREF(Py_None);
			return Py_None;
		}
	case VT_BOOL:
		{
			PyObject* ret = (0 != var->boolVal) ? Py_True : Py_False;
			Py_INCREF(ret);
			return ret;
		}
	case VT_BOOL | VT_BYREF:
		{
			PyObject* ret = (0 != *var->pboolVal) ? Py_True : Py_False;
			Py_INCREF(ret);
			return ret;
		}
	case VT_R4:
		return PyFloat_FromDouble((double) var->fltVal);
	case VT_R4 | VT_BYREF:
		return PyFloat_FromDouble((double) *var->pfltVal);
	case VT_R8:
		return PyFloat_FromDouble(var->dblVal);
	case VT_R8 | VT_BYREF:
		return PyFloat_FromDouble(*var->pdblVal);
	case VT_I2:
		return PyInt_FromLong((long) var->iVal);
	case VT_I2 | VT_BYREF:
		return PyInt_FromLong((long) *var->piVal);
	case VT_I4:
		return PyInt_FromLong(var->lVal);
	case VT_I4 | VT_BYREF:
		return PyInt_FromLong(*var->plVal);
	case VT_BSTR:
		{
			BOOL bUsedDefaultChar;
			int len = (int) SysStringLen(var->bstrVal);
			char* narrow = new char[len+1];
			WideCharToMultiByte(CP_ACP, 0, var->bstrVal, len, narrow, len+1, "?", &bUsedDefaultChar);
			narrow[len] = 0;
			PyObject* ret = PyString_FromStringAndSize(narrow, (Py_ssize_t) len);
			delete[] narrow;
			if(ret == NULL)
				throw PythonException();
			return ret;
		}
	case VT_BSTR | VT_BYREF:
		{
			BOOL bUsedDefaultChar;
			int len = (int) SysStringLen(*var->pbstrVal);
			char* narrow = new char[len+1];
			WideCharToMultiByte(CP_ACP, 0, *var->pbstrVal, len, narrow, len+1, "?", &bUsedDefaultChar);
			narrow[len] = 0;
			PyObject* ret = PyString_FromStringAndSize(narrow, (Py_ssize_t) len);
			delete[] narrow;
			if(ret == NULL)
				throw PythonException();
			return ret;
		}

	case VT_VARIANT | VT_ARRAY:
	case VT_VARIANT | VT_ARRAY | VT_BYREF:
			return VariantToList(var, dims);

	case VT_VARIANT | VT_BYREF:
			return VariantToPy(var->pvarVal, dims);
	
	case VT_UNKNOWN:
	case VT_UNKNOWN | VT_BYREF:
		{
			IUnknown* pUnk = (var->vt & VT_BYREF) ? *var->ppunkVal : var->punkVal;
			IPyObjectImplPtr pPyObject;

			if(NULL != (pPyObject = pUnk))
				return pPyObject->GetNewRef();
			else
				return IUnknownToPyWin32(pUnk);
		}

	case VT_DISPATCH:
	case VT_DISPATCH | VT_BYREF:
		{
			IDispatch* pDisp = (var->vt & VT_BYREF) ? *var->ppdispVal : var->pdispVal;
			IPyObjectImplPtr pPyObject;

			if(NULL != (pPyObject = pDisp))
				return pPyObject->GetNewRef();
			else
			{
				IUnknownPtr pUnk = pDisp;
				return IUnknownToPyWin32(pUnk);
			}
		}

	default:
		throw Exception() << "Cannot convert variant to Python object.";
	}
}

int SafeArrayLength(SAFEARRAY* pSA)
{
	if(pSA->cDims != 1)
		throw Exception() << "Array is not one-dimensional while checking length.";

	return (int) pSA->rgsabound->cElements;
}


PyObject* SafeArrayToTuple(SAFEARRAY* pSA)
{
	LogCtx log("SafeArrayToTuple");

	if(pSA->cDims != 1)
		throw Exception() << "Array is not one-dimensional while converting to Python tuple.";

	int len = (int) pSA->rgsabound->cElements;
	PyNewRef ret(PyTuple_New(len));
	VARIANT* pData;
	{
		AutoSafeArrayAccessData ad(pSA, (void**) &pData);
		for(int k=0; k<len; k++)
		{
			PyTuple_SET_ITEM(ret.ptr, k, VariantToPy(&pData[k]));
		}
	}

	return ret.detach();
}

PyObject* SafeArrayToList(SAFEARRAY* pSA)
{
	LogCtx log("SafeArrayToList");

	if(pSA->cDims != 1)
		throw Exception() << "Array is not one-dimensional while converting to Python list.";

	int len = (int) pSA->rgsabound->cElements;
	PyNewRef ret(PyList_New(len));
	VARIANT* pData;
	{
		AutoSafeArrayAccessData ad(pSA, (void**) &pData);

		for(int k=0; k<len; k++)
			PyList_SET_ITEM(ret.ptr, k, VariantToPy(&pData[k]));
	}

	return ret.detach();
}

PyObject* VariantToTuple(VARIANT* var)
{
	LogCtx log("VariantToTuple");

	if(var->vt != (VT_VARIANT | VT_ARRAY))
		throw Exception() << "Variant is not array of variants while converting to Python tuple.";

	return SafeArrayToTuple(var->parray);
}

PyObject* SafeArrayToDict(SAFEARRAY* pSA)
{
	LogCtx log("SafeArrayToDict");

	switch(pSA->cDims)
	{
	case 1:
		{
			if(pSA->rgsabound[0].cElements % 2 != 0)
				throw Exception() << "Array is not of length 2n while converting to Python dictionary."; 

			int len = (int) pSA->rgsabound[0].cElements / 2;

			PyNewRef ret(PyDict_New());
			VARIANT* pData;
			{
				AutoSafeArrayAccessData ad(pSA, (void**) &pData);

				for(int k=0; k<len; k++)
				{
					PyNewRef key(VariantToPy(&pData[2*k]));
					PyNewRef value(VariantToPy(&pData[2*k+1]));
					PyDict_SetItem(ret.ptr, key, value);
				}
			}

			return ret.detach();
		}
		break;
		
	case 2:
		{
			if(pSA->rgsabound[0].cElements != 2)
				throw Exception() << "Array is not (n x 2) while converting to Python dictionary."; 

			int len = (int) pSA->rgsabound[1].cElements;

			PyNewRef ret(PyDict_New());
			VARIANT* pData;
			{
				AutoSafeArrayAccessData ad(pSA, (void**) &pData);

				for(int k=0; k<len; k++)
				{
					PyNewRef key(VariantToPy(&pData[k]));
					PyNewRef value(VariantToPy(&pData[k+len]));
					PyDict_SetItem(ret.ptr, key, value);
				}
			}

			return ret.detach();
		}
		break;

	default:
		throw Exception() << "Array is not one- or two-dimensional while converting to Python dictionary.";
	}
}

PyObject* VariantToDict(VARIANT* var)
{
	LogCtx log("VariantToDict");

	if(var->vt != (VT_VARIANT | VT_ARRAY))
		throw Exception() << "Variant is not array of variants while converting to Python dictionary.";

	return SafeArrayToDict(var->parray);
}

PyObject* VariantToList(VARIANT* var, int dims)
{
	LogCtx log("VariantToList");

	if(var->vt != (VT_VARIANT | VT_ARRAY) && var->vt != (VT_VARIANT | VT_ARRAY | VT_BYREF))
		throw Exception() << "Variant is not array of variants while converting to Python list.";

	SAFEARRAY* pSA = (var->vt & VT_BYREF) ? *var->pparray : var->parray;
	if(pSA->cDims == 1)
	{
		int len = (int) pSA->rgsabound->cElements;
		PyNewRef ret(PyList_New(len));
		VARIANT* pData;
		{
			AutoSafeArrayAccessData ad(pSA, (void**) &pData);
			for(int k=0; k<len; k++)
				PyList_SET_ITEM(ret.ptr, k, VariantToPy(&pData[k]));
		}
		return ret.detach();
	}
	else if(pSA->cDims == 2)
	{
		if(dims == 1)
		{
			int len1 = (int) pSA->rgsabound[1].cElements;
			int len2 = (int) pSA->rgsabound[0].cElements;
			int count = len1 * len2;
			PyNewRef ret(PyList_New(count));
			if(count > 0)
			{
				if(len1 != 1 && len2 != 1)
					throw Exception() << "To convert 2-dimensional variant array to 1-dimensional list it must be of size 1 x n or n x 1.";

				VARIANT* pData;
				{
					AutoSafeArrayAccessData ad(pSA, (void**) &pData);
					for(int k=0; k<count; k++)
						PyList_SET_ITEM(ret.ptr, k, VariantToPy(&pData[k]));
				}
			}
			return ret.detach();
		}
		else
		{
			int len1 = (int) pSA->rgsabound[1].cElements;
			int len2 = (int) pSA->rgsabound[0].cElements;
			PyNewRef ret(PyList_New(len1));
			VARIANT* pData;
			{
				AutoSafeArrayAccessData ad(pSA, (void**) &pData);
				for(int k=0; k<len1; k++)
				{
					PyNewRef sub(PyList_New(len2));
					for(int j=0; j<len2; j++)
						PyList_SET_ITEM(sub.ptr, j, VariantToPy(&pData[k + j * len1]));
					PyList_SET_ITEM(ret.ptr, k, sub.detach());
				}
			}
			return ret.detach();
		}
	}
	else
		throw Exception() << "Array is not one or two-dimensional while converting to Python list.";
}

void WrapPyToVariant(PyObject* obj, VARIANT* var)
{
	if(obj == NULL)
		throw Exception() << "Internal error: NULL passed to PyWrapToVariant";

	VariantClear(var);
	//var->vt = VT_UNKNOWN;
	//var->punkVal = IPyObjectImpl::FromBorrowedRef(obj);
	var->vt = VT_DISPATCH;
	var->pdispVal = IPyObjectImpl::FromBorrowedRef(obj);
}

void WrapPyToIPyObject(PyObject* obj, IPyObj** pPy)
{
	if(obj == NULL)
		throw Exception() << "Internal error: NULL passed to WrapPyToIPyObject";

	*pPy = IPyObjectImpl::FromBorrowedRef(obj);
}

void ConvertPyToVariant(PyObject* obj, VARIANT* var, bool deep, int dims)
{
	LogCtx log("ConvertPyToVariant");

	VariantClear(var);

	if(obj == NULL || obj == Py_None)
		return;

	if(PyFloat_CheckExact(obj))
	{
		var->vt = VT_R8;
		var->dblVal = PyFloat_AS_DOUBLE(obj);
	}
	else if(PyInt_CheckExact(obj))
	{
		var->vt = VT_I4;
		var->lVal = PyInt_AS_LONG(obj);
	}
	else if(PyBool_Check(obj))
	{
		var->vt = VT_BOOL;
		var->boolVal = (obj == Py_True);
	}
	else if(PyString_CheckExact(obj))
	{
		char* str;
		Py_ssize_t len;
		if(0 != PyString_AsStringAndSize(obj, &str, &len))
			throw PythonException();

		int sz = (int) len + 1;
		OLECHAR* wide = new OLECHAR[sz];
		MultiByteToWideChar(CP_ACP, 0, str, sz * sizeof(OLECHAR), wide, sz);
		var->vt = VT_BSTR;
		var->bstrVal = SysAllocString(wide);
		delete[] wide;
	}
	else if(PyUnicode_CheckExact(obj))
	{
		wchar_t* str = PyUnicode_AS_UNICODE(obj);
		if(str == NULL)
			throw PythonException();
		Py_ssize_t len = PyUnicode_GET_SIZE(obj);
		int sz = (int) len + 1;
		var->vt = VT_BSTR;
		var->bstrVal = SysAllocString(str);
	}
	else if(deep && PyList_CheckExact(obj))
	{
		Py_ssize_t len = PyList_Size(obj);

		if(dims == 2)
		{
			Py_ssize_t len2 = -1;
			boost::scoped_array<PyObject*> elements(new PyObject*[len]);
			if(len == 0)
				len2 = 0;
			else
			{
				for(int k=0; k<len; k++)
				{
					PyObject* e = PyList_GetItem(obj, k);
					if(!PyList_CheckExact(e))
						throw Exception() << "Cannot convert list to 2-dimensional variant array because at least one element isn't a list.";
					else
					{
						if(len2 < 0)
							len2 = PyList_Size(e);
						else if(PyList_Size(e) != len2)
							throw Exception() << "Cannot convert list of lists to 2-dimensional variant array because lengths differ.";
					}
				}
			}

			SAFEARRAYBOUND bounds[2];
			bounds[1].lLbound = 1;
			bounds[1].cElements = (ULONG) len2;
			bounds[0].lLbound = 1;
			bounds[0].cElements = (ULONG) len;

			AutoSafeArrayCreate asac(VT_VARIANT, 2, bounds);

			VARIANT* pData;
			{
				AutoSafeArrayAccessData(asac.pSafeArray, (void**) &pData);

				for(int k=0; k<len*len2; k++)
					VariantInit(&pData[k]);

				for(int k=0; k<len; k++)
				{
					PyObject* e = PyList_GetItem(obj, k);
					for(int j=0; j<len2; j++)
						ConvertPyToVariant(PyList_GetItem(e, j), &pData[k + j * len], deep);
				}
			}

			var->vt = VT_VARIANT | VT_ARRAY;
			var->parray = asac.pSafeArray;
			asac.pSafeArray = NULL;
		}
		else
		{
			SAFEARRAYBOUND bounds[1];
			bounds[0].lLbound = 1;
			bounds[0].cElements = (ULONG) len;
			AutoSafeArrayCreate asac(VT_VARIANT, 1, bounds);

			VARIANT* pData;
			{
				AutoSafeArrayAccessData(asac.pSafeArray, (void**) &pData);
				for(int k=0; k<len; k++)
					VariantInit(&pData[k]);

				for(int k=0; k<len; k++)
					ConvertPyToVariant(PyList_GetItem(obj, k), &pData[k], deep);
			}

			var->vt = VT_VARIANT | VT_ARRAY;
			var->parray = asac.pSafeArray;
			asac.pSafeArray = NULL;
		}
	}
	else
	{
		var->vt = VT_UNKNOWN;
		var->punkVal = IPyObjectImpl::FromBorrowedRef(obj);
	}
}