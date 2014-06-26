#include "ExcelPython.h"

#include <hash_map>

_COM_SMARTPTR_TYPEDEF(ITypeInfo, __uuidof(ITypeInfo));
_COM_SMARTPTR_TYPEDEF(ITypeLib, __uuidof(ITypeLib));

struct PyDispatchTypeObject : PyTypeObject
{
	PyDispatchTypeObject()
	{
		memset(this, 0, sizeof(PyTypeObject));
		this->ob_type = &PyType_Type;
		this->ob_refcnt = 1;
	}
};

stdext::hash_map<std::string, PyDispatchTypeObject*> _dispatch_types;

struct PyDispatch
{
	PyObject_HEAD;
	IDispatch* pDispatch;
};

struct PyDispatchProperty
{
	PyObject_HEAD;
	PyDispatchTypeObject* parent;
	PyStringObject* name;
	MEMBERID idGet;
	MEMBERID idPut;
	MEMBERID idPutRef;
	bool bIndexed;
};

struct PyBoundDispatchProperty
{
	PyObject_HEAD;
	PyDispatchProperty* parent;
	IDispatch* pDispatch;
};

struct PyDispatchMethod
{
	PyObject_HEAD;
	PyDispatchTypeObject* parent;
	MEMBERID id;
};

struct PyBoundDispatchMethod
{
	PyObject_HEAD;
	PyDispatchMethod* parent;
	IDispatch* pDispatch;
};

/***************************** PyDispatchProperty ******************************/

PyTypeObject* PyDispatchProperty_Type();
PyObject* PyDispatchProperty_New(PyDispatchTypeObject* parent, char* name)
{
	PyDispatchProperty* obj = PyObject_New(PyDispatchProperty, PyDispatchProperty_Type());
	obj->bIndexed = false;
	obj->name = (PyStringObject*) PyString_FromString(name);
	obj->idGet = MEMBERID_NIL;
	obj->idPut = MEMBERID_NIL;
	obj->idPutRef = MEMBERID_NIL;
	obj->parent = parent;
	Py_INCREF(parent);
	return (PyObject*) obj;	
}

PyTypeObject* PyBoundDispatchProperty_Type();
PyObject* PyBoundDispatchProperty_New(PyDispatchProperty* parent, IDispatch* pDispatch)
{
	PyBoundDispatchProperty* obj = PyObject_New(PyBoundDispatchProperty, PyBoundDispatchProperty_Type());
	obj->parent = parent;
	Py_INCREF(parent);
	obj->pDispatch = pDispatch;
	pDispatch->AddRef();
	return (PyObject*) obj;	
}

PyObject * PyDispatchProperty_Descr_Get(PyDispatchProperty* self, PyDispatch* target, PyObject* type)
{
	if(self->bIndexed)
		return PyBoundDispatchProperty_New(self, target->pDispatch);
	else
	{
		DISPPARAMS params;
		params.cNamedArgs = 0;
		params.rgdispidNamedArgs = NULL;
		params.cArgs = 0;
		params.rgvarg = NULL;

		VARIANT result;
		VariantInit(&result);

		EXCEPINFO exInfo;
		UINT uArgErr = 0;

		HRESULT hr = target->pDispatch->Invoke(self->idGet, IID_NULL, 0, DISPATCH_PROPERTYGET, &params, &result, &exInfo, &uArgErr);
		if(FAILED(hr))
		{
			_com_error err(hr);
			PyErr_SetString(PyExc_RuntimeError, WCharToStdString(err.ErrorMessage()).c_str());
			return NULL;
		}

		return VariantToPy(&result);
	}
}

int PyBoundDispatchProperty_Length(PyBoundDispatchProperty* self)
{
	return 0;
}

PyObject* PyBoundDispatchProperty_Subscript(PyBoundDispatchProperty* self, PyObject* index)
{
	DISPPARAMS params;
	params.cNamedArgs = 0;
	params.rgdispidNamedArgs = NULL;

	if(Py_TYPE(index) == &PySlice_Type)
		params.cArgs = 0;
	else if(PyTuple_CheckExact(index))
		params.cArgs = (UINT) PyTuple_GET_SIZE(index);
	else
		params.cArgs = 1;

	boost::scoped_array<VARIANT> vargs(new VARIANT[params.cArgs]);
	params.rgvarg = vargs.get();
	
	if(Py_TYPE(index) == &PySlice_Type)
	{
	}
	else if(PyTuple_CheckExact(index))
	{
		for(UINT iArg = 0; iArg < params.cArgs; iArg++)
		{
			VARIANT* varg = vargs.get() + (params.cArgs - iArg - 1); 
			VariantInit(varg);
			ConvertPyToVariant(PyTuple_GET_ITEM(index, iArg), varg, true);
		}
	}
	else
	{
		VariantInit(vargs.get());
		ConvertPyToVariant(index, vargs.get(), true);
	}

	VARIANT result;
	VariantInit(&result);

	EXCEPINFO exInfo;
	UINT uArgErr = 0;

	HRESULT hr = self->pDispatch->Invoke(self->parent->idGet, IID_NULL, 0, DISPATCH_PROPERTYGET, &params, &result, &exInfo, &uArgErr);
	if(FAILED(hr))
	{
		_com_error err(hr);
		PyErr_SetString(PyExc_RuntimeError, WCharToStdString(err.ErrorMessage()).c_str());
		return NULL;
	}

	return VariantToPy(&result);	
}

int PyBoundDispatchProperty_AssignSubscript(PyBoundDispatchProperty* self, PyObject* index, PyObject* value)
{
	DISPPARAMS params;
	params.cNamedArgs = 0;
	params.rgdispidNamedArgs = NULL;

	if(PyTuple_CheckExact(index))
		params.cArgs = (UINT) PyTuple_GET_SIZE(index) + 1;
	else
		params.cArgs = 2;

	boost::scoped_array<VARIANT> vargs(new VARIANT[params.cArgs]);
	params.rgvarg = vargs.get();
	
	if(PyTuple_CheckExact(index))
	{
		for(UINT iArg = 0; iArg < params.cArgs; iArg++)
		{
			VARIANT* varg = vargs.get() + (params.cArgs - iArg - 1); 
			VariantInit(varg);
			ConvertPyToVariant(PyTuple_GET_ITEM(index, iArg), varg, true);
		}
		VariantInit(vargs.get() + params.cArgs - 1);
		ConvertPyToVariant(value, vargs.get() + params.cArgs - 1, true);
	}
	else
	{
		VariantInit(vargs.get());
		ConvertPyToVariant(index, vargs.get(), true);
		VariantInit(vargs.get() + 1);
		ConvertPyToVariant(value, vargs.get() + 1, true);
	}

	EXCEPINFO exInfo;
	UINT uArgErr = 0;

	HRESULT hr = self->pDispatch->Invoke(self->parent->idPut, IID_NULL, 0, DISPATCH_PROPERTYPUT, &params, NULL, &exInfo, &uArgErr);
	if(FAILED(hr))
	{
		_com_error err(hr);
		PyErr_SetString(PyExc_RuntimeError, WCharToStdString(err.ErrorMessage()).c_str());
		return NULL;
	}

	return 0;
}

PyMappingMethods PyBoundDispatchProperty_MappingMethods = 
{
	(lenfunc) PyBoundDispatchProperty_Length,
	(binaryfunc) PyBoundDispatchProperty_Subscript,
	(objobjargproc) PyBoundDispatchProperty_AssignSubscript
};

int PyDispatchProperty_Descr_Set(PyDispatchProperty* self, PyDispatch* target, PyObject* value)
{
	DISPPARAMS params;
	params.cNamedArgs = 1;
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	params.rgdispidNamedArgs = &dispidNamed;

	params.cArgs = 1;

	boost::scoped_array<VARIANT> vargs(new VARIANT[params.cArgs]);
	params.rgvarg = vargs.get();
	
	VariantInit(vargs.get());
	ConvertPyToVariant(value, vargs.get(), true);

	EXCEPINFO exInfo;
	UINT uArgErr = 0;

	HRESULT hr = target->pDispatch->Invoke(self->idPut, IID_NULL, 0, DISPATCH_PROPERTYPUT, &params, NULL, &exInfo, &uArgErr);
	if(FAILED(hr))
	{
		_com_error err(hr);
		PyErr_SetString(PyExc_RuntimeError, WCharToStdString(err.ErrorMessage()).c_str());
		return -1;
	}

	return 0;
}

void PyBoundDispatchProperty_Dealloc(PyBoundDispatchProperty* self)
{
	self->pDispatch->Release();
	Py_DECREF(self->parent);
}

PyTypeObject* PyDispatchProperty_Type()
{
	static bool ready = false;
	static PyTypeObject type;
	if(!ready)
	{
		memset(&type, 0, sizeof(PyTypeObject));
		type.ob_type = &PyType_Type;
		type.ob_refcnt = 1;
		type.tp_name = "IDispatch property";
		type.tp_basicsize = sizeof(PyDispatchProperty);
		type.tp_flags = Py_TPFLAGS_DEFAULT;
		type.tp_doc = "IDispatch property";
		type.tp_descr_get = (descrgetfunc) PyDispatchProperty_Descr_Get;
		type.tp_descr_set = (descrsetfunc) PyDispatchProperty_Descr_Set;

		if(0 != PyType_Ready(&type))
			throw PythonException();

		ready = true;
	}

	return &type;
}

PyTypeObject* PyBoundDispatchProperty_Type()
{
	static bool ready = false;
	static PyTypeObject type;
	if(!ready)
	{
		memset(&type, 0, sizeof(PyTypeObject));
		type.ob_type = &PyType_Type;
		type.ob_refcnt = 1;
		type.tp_name = "Bound IDispatch property";
		type.tp_basicsize = sizeof(PyBoundDispatchProperty);
		type.tp_dealloc = (destructor) PyBoundDispatchProperty_Dealloc;
		type.tp_flags = Py_TPFLAGS_DEFAULT;
		type.tp_doc = "Bound IDispatch property";
		type.tp_as_mapping = &PyBoundDispatchProperty_MappingMethods;

		if(0 != PyType_Ready(&type))
			throw PythonException();

		ready = true;
	}

	return &type;
}


/***************************** PyDispatchMethod ******************************/

PyTypeObject* PyDispatchMethod_Type();
PyObject* PyDispatchMethod_New(PyDispatchTypeObject* parent, MEMBERID id)
{
	PyDispatchMethod* obj = PyObject_New(PyDispatchMethod, PyDispatchMethod_Type());
	obj->parent = parent;
	obj->id = id;
	return (PyObject*) obj;	
}

PyTypeObject* PyBoundDispatchMethod_Type();
PyObject* PyBoundDispatchMethod_New(PyDispatchMethod* parent, IDispatch* target = NULL)
{
	PyBoundDispatchMethod* obj = PyObject_New(PyBoundDispatchMethod, PyBoundDispatchMethod_Type());
	obj->parent = parent;
	Py_INCREF(parent);
	obj->pDispatch = target;
	target->AddRef();
	return (PyObject*) obj;	
}

PyObject* PyBoundDispatchMethod_Call(PyBoundDispatchMethod* self, PyTupleObject* args, PyDictObject* kwargs)
{
	if(self->pDispatch == NULL)
	{
		PyErr_SetString(PyExc_RuntimeError, "Cannot call unbound IDispatch method.");
		return NULL;
	}

	DISPPARAMS params;
	params.cNamedArgs = 0;
	params.rgdispidNamedArgs = NULL;
	params.cArgs = (UINT) PyTuple_GET_SIZE(args);
	boost::scoped_array<VARIANT> vargs(new VARIANT[params.cArgs]);
	params.rgvarg = vargs.get();
	for(UINT iArg = 0; iArg < params.cArgs; iArg++)
	{
		VARIANT* varg = vargs.get() + (params.cArgs - iArg - 1); 
		VariantInit(varg);
		ConvertPyToVariant(PyTuple_GET_ITEM(args, iArg), varg, true);
	}

	VARIANT result;
	VariantInit(&result);

	EXCEPINFO exInfo;
	UINT uArgErr = 0;

	HRESULT hr = self->pDispatch->Invoke(self->parent->id, IID_NULL, 0, DISPATCH_METHOD, &params, &result, &exInfo, &uArgErr);
	if(FAILED(hr))
	{
		_com_error err(hr);
		PyErr_SetString(PyExc_RuntimeError, WCharToStdString(err.ErrorMessage()).c_str());
		return NULL;
	}

	return VariantToPy(&result);
}

static PyObject* PyDispatchMethod_Descr_Get(PyDispatchMethod* self, PyDispatch* target, PyObject *cls)
{
	return PyBoundDispatchMethod_New(self, target->pDispatch);
}

void PyBoundDispatchMethod_Dealloc(PyBoundDispatchMethod* self)
{
	Py_DECREF(self->parent);
	self->pDispatch->Release();
}

PyTypeObject* PyDispatchMethod_Type()
{
	static bool ready = false;
	static PyTypeObject type;
	if(!ready)
	{
		memset(&type, 0, sizeof(PyTypeObject));
		type.ob_type = &PyType_Type;
		type.ob_refcnt = 1;
		type.tp_name = "Unbound IDispatch method";
		type.tp_basicsize = sizeof(PyDispatchMethod);
		type.tp_flags = Py_TPFLAGS_DEFAULT;
		type.tp_doc = "Unbound IDispatch method";
		type.tp_descr_get = (descrgetfunc) PyDispatchMethod_Descr_Get;

		if(0 != PyType_Ready(&type))
			throw PythonException();

		ready = true;
	}

	return &type;
}

PyTypeObject* PyBoundDispatchMethod_Type()
{
	static bool ready = false;
	static PyTypeObject type;
	if(!ready)
	{
		memset(&type, 0, sizeof(PyTypeObject));
		type.ob_type = &PyType_Type;
		type.ob_refcnt = 1;
		type.tp_name = "Bound IDispatch method";
		type.tp_basicsize = sizeof(PyDispatchMethod);
		type.tp_dealloc = (destructor) PyBoundDispatchMethod_Dealloc;
		type.tp_call = (ternaryfunc) PyBoundDispatchMethod_Call;
		type.tp_flags = Py_TPFLAGS_DEFAULT;
		type.tp_doc = "Bound IDispatch method";

		if(0 != PyType_Ready(&type))
			throw PythonException();

		ready = true;
	}

	return &type;
}


class TYPEATTRPtr
{
public:
	TYPEATTR* _pTypeAttr;
	ITypeInfo* _pTypeInfo;

	TYPEATTRPtr(ITypeInfo* pTypeInfo)
		: _pTypeInfo(pTypeInfo)
	{
		_pTypeInfo->GetTypeAttr(&_pTypeAttr);
	}

	~TYPEATTRPtr()
	{
		_pTypeInfo->ReleaseTypeAttr(_pTypeAttr);
	}

	TYPEATTR* operator->()
	{
		return _pTypeAttr;
	}
};

class FUNCDESCPtr
{
public:
	FUNCDESC* _pFuncDesc;
	ITypeInfo* _pTypeInfo;

	FUNCDESCPtr(ITypeInfo* pTypeInfo, int i)
		: _pTypeInfo(pTypeInfo)
	{
		_pTypeInfo->GetFuncDesc(i, &_pFuncDesc);
	}

	~FUNCDESCPtr()
	{
		_pTypeInfo->ReleaseFuncDesc(_pFuncDesc);
	}

	FUNCDESC* operator->()
	{
		return _pFuncDesc;
	}
};

void GuidToString(GUID& guid, std::string& str)
{
	RPC_CSTR rpcstr;
	UuidToStringA(&guid, &rpcstr);
	str = (char*) rpcstr;
	RpcStringFreeA(&rpcstr);
}

void PyDispatch_Dealloc(PyDispatch* self)
{
	self->pDispatch->Release();
}

Py_ssize_t PyDispatch_Length(PyDispatch* self)
{
	return 0;
}

PyObject* PyDispatch_Subscript(PyDispatch* self, PyObject* index)
{
	DISPPARAMS params;
	params.cNamedArgs = 0;
	params.rgdispidNamedArgs = NULL;

	if(PyTuple_CheckExact(index))
		params.cArgs = (UINT) PyTuple_GET_SIZE(index);
	else
		params.cArgs = 1;

	boost::scoped_array<VARIANT> vargs(new VARIANT[params.cArgs]);
	params.rgvarg = vargs.get();
	
	if(PyTuple_CheckExact(index))
	{
		for(UINT iArg = 0; iArg < params.cArgs; iArg++)
		{
			VARIANT* varg = vargs.get() + (params.cArgs - iArg - 1); 
			VariantInit(varg);
			ConvertPyToVariant(PyTuple_GET_ITEM(index, iArg), varg, true);
		}
	}
	else
	{
		VariantInit(vargs.get());
		ConvertPyToVariant(index, vargs.get(), true);
	}

	VARIANT result;
	VariantInit(&result);

	EXCEPINFO exInfo;
	UINT uArgErr = 0;

	HRESULT hr = self->pDispatch->Invoke(DISPID_VALUE, IID_NULL, 0, DISPATCH_PROPERTYGET, &params, &result, &exInfo, &uArgErr);
	if(FAILED(hr))
	{
		PyErr_SetString(PyExc_RuntimeError, BStrToStdString(exInfo.bstrDescription).c_str());
		return NULL;
	}

	return VariantToPy(&result);	
}

int PyDispatch_AssignSubscript(PyDispatch* self, PyObject* index, PyObject* value)
{
	DISPPARAMS params;
	params.cNamedArgs = 1;
	DISPID dispid = DISPID_PROPERTYPUT;
	params.rgdispidNamedArgs = &dispid;

	if(PyTuple_CheckExact(index))
		params.cArgs = (UINT) PyTuple_GET_SIZE(index) + 1;
	else
		params.cArgs = 2;

	boost::scoped_array<VARIANT> vargs(new VARIANT[params.cArgs]);
	params.rgvarg = vargs.get();
	
	if(PyTuple_CheckExact(index))
	{
		for(UINT iArg = 0; iArg < params.cArgs; iArg++)
		{
			VARIANT* varg = vargs.get() + (params.cArgs - iArg - 1); 
			VariantInit(varg);
			ConvertPyToVariant(PyTuple_GET_ITEM(index, iArg), varg, true);
		}
		VariantInit(vargs.get());
		ConvertPyToVariant(value, vargs.get(), true);
	}
	else
	{
		VariantInit(vargs.get() + 1);
		ConvertPyToVariant(index, vargs.get() + 1, true);
		VariantInit(vargs.get());
		ConvertPyToVariant(value, vargs.get(), true);
	}

	EXCEPINFO exInfo;
	UINT uArgErr = 0;

	HRESULT hr = self->pDispatch->Invoke(DISPID_VALUE, IID_NULL, 0, DISPATCH_PROPERTYPUT, &params, NULL, &exInfo, &uArgErr);
	if(FAILED(hr))
	{
		PyErr_SetString(PyExc_RuntimeError, BStrToStdString(exInfo.bstrDescription).c_str());
		return NULL;
	}

	return 0;
}

PyMappingMethods PyDispatch_MappingMethods = 
{
	(lenfunc) PyDispatch_Length,
	(binaryfunc) PyDispatch_Subscript,
	(objobjargproc) PyDispatch_AssignSubscript
};

PyObject* PyDispatch_New(IDispatch* pDispatch)
{
	HRESULT hr;

	ITypeInfoPtr pTypeInfo;
	if(FAILED(hr = pDispatch->GetTypeInfo(0, 0, &pTypeInfo)))
		throw ComException(hr);

	TYPEATTRPtr pTA(pTypeInfo);
	std::string strGuid;
	GuidToString(pTA->guid, strGuid);
	stdext::hash_map<std::string, PyDispatchTypeObject*>::iterator it = _dispatch_types.find(strGuid);
	if(it != _dispatch_types.end())
	{
		PyDispatch* obj = PyObject_New(PyDispatch, it->second);
		obj->pDispatch = pDispatch;
		pDispatch->AddRef();
		return (PyObject*) obj;	
	}
	else
	{
		ITypeLibPtr pTypeLib;
		if(FAILED(hr = pTypeInfo->GetContainingTypeLib(&pTypeLib, NULL)))
			throw ComException(hr);

		_bstr_t strLibName;
		if(FAILED(hr = pTypeLib->GetDocumentation(MEMBERID_NIL, strLibName.GetAddress(), NULL, NULL, NULL)))
			throw ComException(hr);

		_bstr_t strName;
		if(FAILED(hr = pTypeInfo->GetDocumentation(MEMBERID_NIL, strName.GetAddress(), NULL, NULL, NULL)))
			throw ComException(hr);

		char* strFullName = new char[strLibName.length() + 1 + strName.length() + 1];
		WideCharToMultiByte(CP_ACP, 0, strLibName, strLibName.length(), strFullName, strLibName.length() + 1, "?", NULL);
		WideCharToMultiByte(CP_ACP, 0, strName, strName.length(), strFullName+strLibName.length()+1, strName.length() + 1, "?", NULL);
		strFullName[strLibName.length()] = '.';
		strFullName[strLibName.length() + strName.length() + 1] = 0;

		PyDispatchTypeObject* type = new PyDispatchTypeObject();
		type->ob_refcnt = 1;
		type->tp_basicsize = sizeof(PyDispatch);
		type->tp_name = strFullName;
		type->tp_flags = Py_TPFLAGS_DEFAULT | Py_TPFLAGS_BASETYPE;
		type->tp_doc = strFullName;
		type->tp_as_mapping = &PyDispatch_MappingMethods;
		type->tp_dealloc = (destructor) PyDispatch_Dealloc;

		PyNewRef attrs(PyDict_New());
		for(int iFunc=0; iFunc < pTA->cFuncs; iFunc++)
		{
			FUNCDESCPtr pFunc(pTypeInfo, iFunc);
			UINT cNames = 1;
			boost::scoped_array<BSTR> pNames(new BSTR[1]);
			if(FAILED(hr = pTypeInfo->GetNames(pFunc->memid, pNames.get(), cNames, &cNames)))
				throw ComException(hr);

			BOOL bUsedDefaultChar;
			int len = (int) SysStringLen(pNames[0]);
			boost::scoped_array<char> narrow(new char[len+1]);
			WideCharToMultiByte(CP_ACP, 0, pNames[0], len, narrow.get(), len+1, "?", &bUsedDefaultChar);
			narrow[len] = 0;
			switch(pFunc->invkind)
			{
			case INVOKE_FUNC:
				{
					PyNewRef method(PyDispatchMethod_New(type, pFunc->memid));
					PyDict_SetItemString(attrs, narrow.get(), method);
					break;
				}

			case INVOKE_PROPERTYGET:
				{
					PyRef prop = PyBorrowedRef(PyDict_GetItemString(attrs, narrow.get()), false);
					if(prop == NULL)
					{
						prop = PyNewRef(PyDispatchProperty_New(type, narrow.get()));
						PyDict_SetItemString(attrs, narrow.get(), prop);
					}
					((PyDispatchProperty*) prop.ptr)->bIndexed = pFunc->cParams > pFunc->cParamsOpt;
					//((PyDispatchProperty*) prop.ptr)->bIndexed = ((PyDispatchProperty*) prop.ptr)->bIndexed || (pFunc->cParams > pFunc->cParamsOpt);
					((PyDispatchProperty*) prop.ptr)->idGet = pFunc->memid;
					break;
				}

			case INVOKE_PROPERTYPUT:
				{
					PyRef prop = PyBorrowedRef(PyDict_GetItemString(attrs, narrow.get()), false);
					if(prop == NULL)
					{
						prop = PyNewRef(PyDispatchProperty_New(type, narrow.get()));
						PyDict_SetItemString(attrs, narrow.get(), prop);
					}
					//((PyDispatchProperty*) prop.ptr)->bIndexed = ((PyDispatchProperty*) prop.ptr)->bIndexed || (pFunc->cParams > pFunc->cParamsOpt);
					((PyDispatchProperty*) prop.ptr)->idPut = pFunc->memid;
					break;
				}

			case INVOKE_PROPERTYPUTREF:
				{
					PyRef prop = PyBorrowedRef(PyDict_GetItemString(attrs, narrow.get()), false);
					if(prop == NULL)
					{
						prop = PyNewRef(PyDispatchProperty_New(type, narrow.get()));
						PyDict_SetItemString(attrs, narrow.get(), prop);
					}
					//((PyDispatchProperty*) prop.ptr)->bIndexed = ((PyDispatchProperty*) prop.ptr)->bIndexed || (pFunc->cParams > pFunc->cParamsOpt);
					((PyDispatchProperty*) prop.ptr)->idPutRef = pFunc->memid;
					break;
				}

			default:
				throw Exception() << "Not implemented.";
			}
		}
		type->tp_dict = attrs.detach();

		if(0 != PyType_Ready(type))
			throw PythonException();

		_dispatch_types[strGuid] = type;

		PyDispatch* obj = PyObject_New(PyDispatch, type);
		obj->pDispatch = pDispatch;
		pDispatch->AddRef();
		return (PyObject*) obj;	
	}
}


//
///******************* IStatusCallback ******************/
//
//struct PyProgressCallback
//{
//	PyObject_HEAD;
//	IProgressCallback* pCallback;
//};
//
//PyTypeObject* PyProgressCallback_Type();
//
//PyObject* PyProgressCallback_New(IProgressCallback* pCallback)
//{
//	PyProgressCallback* obj = PyObject_New(PyProgressCallback, PyProgressCallback_Type());
//	obj->pCallback = pCallback;
//	obj->pCallback->AddRef();
//	return (PyObject*) obj;	
//}
//
//void PyProgressCallback_Dealloc(PyProgressCallback* self)
//{
//	self->pCallback->Release();
//}
//
//PyObject* PyProgressCallback_Call(PyProgressCallback* self, PyTupleObject* args, PyDictObject* kwargs)
//{
//	if(PyTuple_GET_SIZE(args) != 1)
//	{
//		PyErr_SetString(PyExc_RuntimeError, "Expected 1 argument.");
//		return NULL;
//	}
//
//	PyBorrowedRef objStatus(PyTuple_GET_ITEM(args, 0), false);
//	if(!PyString_Check(objStatus))
//	{
//		PyErr_SetString(PyExc_RuntimeError, "Expected string argument.");
//		return NULL;
//	}
//
//	BSTR bstrStatus = CStrToBStr(PyString_AsString(objStatus));
//	HRESULT hr = self->pCallback->SetStatus(&bstrStatus);
//	SysFreeString(bstrStatus);
//
//	if(FAILED(hr))
//	{
//		PyErr_SetString(PyExc_RuntimeError, "COM error.");
//		return NULL;
//	}
//
//	Py_RETURN_NONE;
//}
//
//PyTypeObject* PyProgressCallback_Type()
//{
//	static bool ready = false;
//	static PyTypeObject type;
//	if(!ready)
//	{
//		memset(&type, 0, sizeof(PyTypeObject));
//		type.ob_type = &PyType_Type;
//		type.ob_refcnt = 1;
//		type.tp_name = "Progress Callback";
//		type.tp_basicsize = sizeof(PyProgressCallback);
//		type.tp_flags = Py_TPFLAGS_DEFAULT;
//		type.tp_doc = "Progress Callback";
//		type.tp_call = (ternaryfunc) PyProgressCallback_Call;
//		type.tp_dealloc = (destructor) PyProgressCallback_Dealloc;
//
//		if(0 != PyType_Ready(&type))
//			throw PythonException();
//
//		ready = true;
//	}
//
//	return &type;
//}