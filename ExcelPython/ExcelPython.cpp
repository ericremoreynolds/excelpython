#include "ExcelPython.h"

std::string strDefaultPythonPath;

HINSTANCE hModuleDLL;

BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	hModuleDLL = hinstDLL;

	return TRUE;
}

class ExecutionSentinel
{
	static HANDLE GetActCtxHandle()
	{
		static HANDLE hActCtx = INVALID_HANDLE_VALUE;
		if(hActCtx == INVALID_HANDLE_VALUE)
		{
			LogCtx log("ExecutionSentinel::GetActCtxHandle");

			TCHAR moduleName[MAX_PATH];
			if(0 == GetModuleFileName(hModuleDLL, moduleName, MAX_PATH))
				throw Exception() << "GetModuleFileName failed in ExecutionSentinel::GetActCtxHandle";

			ACTCTX actctx;
			SecureZeroMemory(&actctx, sizeof(actctx));
			actctx.cbSize = sizeof(actctx);
			actctx.lpSource = moduleName;
			actctx.dwFlags = ACTCTX_FLAG_HMODULE_VALID | ACTCTX_FLAG_RESOURCE_NAME_VALID;
			actctx.hModule = hModuleDLL;
			actctx.lpResourceName = MAKEINTRESOURCE(ISOLATIONAWARE_MANIFEST_RESOURCE_ID);
			hActCtx = CreateActCtx(&actctx);

			if(hActCtx == INVALID_HANDLE_VALUE)
				throw Exception() << "Could not create DLL activation context.";
		}
		return hActCtx;
	}

	static void EnsurePythonInitialized()
	{
		static bool initialized = false;

		if(!initialized)
		{
			LogCtx log("ExecutionSentinel::EnsurePythonInitialized");

			char filename[MAX_PATH];
			GetModuleFileNameA(NULL, filename, MAX_PATH);
			log % "GetModuleFileNameA returned " << filename << ", calling Py_SetProgramName\n";
			Py_SetProgramName(filename);
			char* x = Py_GetPythonHome();
			log % "Py_GetPythonHome returned " << x << ", calling Py_Initialize\n";
			Py_Initialize();
			strDefaultPythonPath = Py_GetPath();
			log % "Py_GetPath returned " << strDefaultPythonPath << "\n";

			{
				// get the default python path
				PyBorrowedRef sys_path(PySys_GetObject("path"));
				PyNewRef semicolon(PyString_FromString(";"));
				PyNewRef str_path(PyObject_CallMethod(semicolon, "join", "(O)", sys_path));
				strDefaultPythonPath = PyString_AsString(str_path);

				log % "sys.path = " << strDefaultPythonPath << "\n";
			}

			PyNewRef console(PyConsoleReadWrite_New());
			console.ptr->ob_refcnt += 3;
			PySys_SetObject("stdout", console);
			PySys_SetObject("stdin", console);
			PySys_SetObject("stderr", console);
			log % "Console set up\n";

			initialized = true;
		}
	}

	ULONG_PTR ulpCookie;

public:
	ExecutionSentinel()
	{
		if(!ActivateActCtx(GetActCtxHandle(), &this->ulpCookie))
			throw Exception() << "Could not activate DLL activation context.";

		EnsurePythonInitialized();
	}

	~ExecutionSentinel()
	{
		if(!DeactivateActCtx(0, this->ulpCookie))
			throw Exception() << "Could not deactivate DLL activation context.";
	}
};

void ShowConsole(bool show)
{
	LogCtx log("ShowConsole");

	if(show)
	{
		AllocConsole();
		HWND hWnd = GetConsoleWindow();
		EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
		BringWindowToTop(hWnd);

		FILE* unused;
		freopen_s(&unused, "conout$", "a", stdout);
		freopen_s(&unused, "conout$", "a", stderr);
		freopen_s(&unused, "conin$", "r", stdin);
		setvbuf(stdout, NULL, _IONBF, 0);
		setvbuf(stderr, NULL, _IONBF, 0);
		setvbuf(stdin, NULL, _IONBF, 0);
	}
	else
	{
		FreeConsole();
	}
}

HRESULT RaiseError(const std::string& msg)
{
	LogCtx log("RaiseError");

	ICreateErrorInfoPtr pCreateErrInfo;
	CreateErrorInfo(&pCreateErrInfo);

	int sz = (int) msg.length() + 1;
	OLECHAR* wide = new OLECHAR[sz];
	MultiByteToWideChar(CP_ACP, 0, msg.c_str(), sz * sizeof(OLECHAR), wide, sz);
	BSTR bstr = SysAllocString(wide);
	delete[] wide;

	pCreateErrInfo->SetDescription(bstr);
	
	IErrorInfoPtr pErrInfo = pCreateErrInfo;
	SetErrorInfo(0, pErrInfo);

	return DISP_E_EXCEPTION;
}

inline bool IsMissing(VARIANT* var)
{
	return var->vt == VT_ERROR && var->scode == DISP_E_PARAMNOTFOUND;
}

void SetupPath(BSTR xlAddPath, BSTR xlPythonPath)
{
	LogCtx log("SetupPath");

	std::string strPath;
	if(0 != SysStringLen(xlPythonPath))
		strPath = BStrToStdString(xlPythonPath);
	else
		strPath = strDefaultPythonPath;

	if(0 != SysStringLen(xlAddPath))
		strPath = BStrToStdString(xlAddPath) + ";" + strPath;

	PySys_SetPath(const_cast<char*>(strPath.c_str()));
}

HRESULT __stdcall PyEval(BSTR xlExpression, VARIANT* xlLocals, VARIANT* xlGlobals, BSTR xlAddPath, BSTR xlPythonPath, IPyObj** xlResult)
{
	LogCtx log("PyEval");

	try
	{

		ExecutionSentinel exe;

		SetupPath(xlAddPath, xlPythonPath);

		PyRef locals, globals;
		if(IsMissing(xlLocals))
			locals = PyBorrowedRef(PyModule_GetDict(PyImport_AddModule("__main__")));
		else
			locals = PyNewRef(VariantToPy(xlLocals));
		if(IsMissing(xlGlobals))
			globals = PyBorrowedRef(PyModule_GetDict(PyImport_AddModule("__main__")));
		else
			globals = PyNewRef(VariantToPy(xlGlobals));

		std::string strExpression = BStrToStdString(xlExpression);
		PyNewRef ret(PyRun_String(strExpression.c_str(), Py_eval_input, globals, locals));

		//VariantInit(xlResult);
		//WrapPyToVariant(ret, xlResult);
		WrapPyToIPyObject(ret, xlResult);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyExec(BSTR xlStatement, VARIANT* xlLocals, VARIANT* xlGlobals, BSTR xlAddPath, BSTR xlPythonPath)
{
	LogCtx log("PyExec");

	try
	{
		ExecutionSentinel exe;

		SetupPath(xlAddPath, xlPythonPath);

		PyRef locals, globals;
		if(IsMissing(xlLocals))
			locals = PyBorrowedRef(PyModule_GetDict(PyImport_AddModule("__main__")));
		else
			locals = PyNewRef(VariantToPy(xlLocals));
		if(IsMissing(xlGlobals))
			globals = PyBorrowedRef(PyModule_GetDict(PyImport_AddModule("__main__")));
		else
			globals = PyNewRef(VariantToPy(xlGlobals));

		std::string strStatement = BStrToStdString(xlStatement);
		PyNewRef ret(PyRun_String(strStatement.c_str(), Py_single_input, globals, locals));

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyTuple(SAFEARRAY** xlElements, VARIANT* xlResult)
{
	LogCtx log("PyTuple");

	try
	{
		ExecutionSentinel exe;

		PyNewRef tuple(SafeArrayToTuple(*xlElements));

		VariantInit(xlResult);
		WrapPyToVariant(tuple, xlResult);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyList(SAFEARRAY** xlElements, VARIANT* xlResult)
{
	LogCtx log("PyList");

	try
	{
		ExecutionSentinel exe;

		PyNewRef list(SafeArrayToList(*xlElements));

		VariantInit(xlResult);
		WrapPyToVariant(list, xlResult);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyDict(SAFEARRAY** xlKeyValuePairs, VARIANT* xlResult)
{
	LogCtx log("PyDict");

	try
	{
		ExecutionSentinel exe;

		PyNewRef dict(SafeArrayToDict(*xlKeyValuePairs));
		
		VariantInit(xlResult);
		WrapPyToVariant(dict, xlResult);
		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyIsObject(VARIANT* xlObject, VARIANT_BOOL* xlResult)
{
	try
	{
		ExecutionSentinel exe;

		switch(xlObject->vt)
		{
			case VT_UNKNOWN:
			case VT_UNKNOWN | VT_BYREF:
			{
				IUnknown* pUnk = (xlObject->vt & VT_BYREF) ? *xlObject->ppunkVal : xlObject->punkVal;
				IPyObjectImplPtr pPyObject = pUnk;

				*xlResult = (NULL != pPyObject) ? -1 : 0;
				break;
			}

			default:
				*xlResult = false;
				break;
		}

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyToObject(VARIANT* xlObject, LONG* xlDimensions, VARIANT* xlResult)
{
	LogCtx log("PyToObject");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlObject, *xlDimensions));

		VariantInit(xlResult);
		WrapPyToVariant(obj, xlResult);
		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}


HRESULT __stdcall PyToVariant(VARIANT* xlObject, LONG* xlDimensions, VARIANT* xlResult)
{
	LogCtx log("PyToVariant");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlObject));

		VariantInit(xlResult);
		ConvertPyToVariant(obj, xlResult, true, *xlDimensions);
		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall IPyObjectImpl::Var(LONG* xlDimensions, VARIANT* xlResult)
{
	try
	{
		VariantInit(xlResult);
		ConvertPyToVariant(this->pObj, xlResult, true, *xlDimensions);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}


//HRESULT __stdcall PySetPath(BSTR xlPythonPath)
//{
//	try
//	{
//		std::string strPythonPath = BStrToStdString(xlPythonPath);
//		PySys_SetPath(const_cast<char*>(strPythonPath.c_str()));
//
//		return S_OK;
//	}
//	catch(Exception& e)
//	{
//		return RaiseError(e.message());
//	}
//}

HRESULT __stdcall PyBuiltin(BSTR xlMethod, VARIANT* xlArgs, VARIANT* xlKeywordArgs, VARIANT* xlResult)
{
	LogCtx log("PyBuiltin");

	try
	{
		ExecutionSentinel exe;

		std::string strMethod = BStrToStdString(xlMethod);

		PyNewRef obj(PyImport_AddModule("__builtin__"));
		PyNewRef method(PyObject_GetAttrString(obj, strMethod.c_str()));

		PyRef args, kwargs;
		if(IsMissing(xlArgs))
			args = PyNewRef(PyTuple_New(0));
		else
		{
			args = PyNewRef(VariantToPy(xlArgs));
			if(!PyTuple_Check(args))
				throw Exception() << "Args should be a tuple object.";
		}
		if(!IsMissing(xlKeywordArgs))
		{
			kwargs = PyNewRef(VariantToPy(xlKeywordArgs));
			if(!PyDict_Check(kwargs))
				throw Exception() << "KwArgs should be a dict object.";
		}

		PyNewRef ret(PyObject_Call(method, args, kwargs));

		VariantInit(xlResult);
		WrapPyToVariant(ret, xlResult);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyStr(VARIANT* xlObject, SAFEARRAY** xlFormatArgs, VARIANT* xlResult)
{
	LogCtx log("PyStr");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlObject));

		PyRef ret = PyNewRef(PyObject_Str(obj));

		// if format args were passed, format the string
		if(SafeArrayLength(*xlFormatArgs) > 0)
		{
			PyNewRef formatArgs(SafeArrayToTuple(*xlFormatArgs));
			ret = PyNewRef(PyString_Format(ret, formatArgs));
		}

		VariantInit(xlResult);
		ConvertPyToVariant(ret, xlResult);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall IPyObjectImpl::Str(VARIANT* xlResult)
{
	try
	{
		ExecutionSentinel exe;

		PyNewRef ret(PyObject_Str(this->pObj));

		VariantInit(xlResult);
		ConvertPyToVariant(ret, xlResult);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyRepr(VARIANT* xlObject, VARIANT* xlResult)
{
	LogCtx log("PyRepr");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlObject));

		PyNewRef ret(PyObject_Repr(obj));

		VariantInit(xlResult);
		ConvertPyToVariant(ret, xlResult);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyLen(VARIANT* xlObject, int* xlResult)
{
	LogCtx log("PyLen");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlObject));

		*xlResult = (int) PyObject_Length(obj);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall IPyObjectImpl::get_Item(VARIANT* xlKey, IPyObj** xlResult)
{
	try
	{
		PyNewRef key(VariantToPy(xlKey));

		PyNewRef ret(PyObject_GetItem(this->pObj, key));

		WrapPyToIPyObject(ret, xlResult);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall IPyObjectImpl::get_Attr(BSTR xlAttribute, IPyObj** xlResult)
{
	try
	{
		std::string strAttr = BStrToStdString(xlAttribute);
		PyNewRef ret(PyObject_GetAttrString(this->pObj, strAttr.c_str()));

		WrapPyToIPyObject(ret, xlResult);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall IPyObjectImpl::get_Count(LONG* xlResult)
{
	try
	{
		*xlResult = (int) PyObject_Length(this->pObj);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}


HRESULT __stdcall PyType(VARIANT* xlObject, VARIANT* xlResult)
{
	LogCtx log("PyType");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlObject));

		PyNewRef ret(PyObject_Type(obj));

		VariantInit(xlResult);
		WrapPyToVariant(ret, xlResult);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyIter(VARIANT* xlInstance, VARIANT* xlResult)
{
	LogCtx log("PyIter");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlInstance));
		PyNewRef ret(PyObject_GetIter(obj));

		VariantInit(xlResult);
		WrapPyToVariant(ret, xlResult);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall IPyObjectImpl::get__NewEnum(IUnknown** ppUnknown)
{
	try
	{
		PyNewRef ret(PyObject_GetIter(this->pObj));

		IPyObj* pPyObj;
		WrapPyToIPyObject(ret, &pPyObj);
		*ppUnknown = pPyObj;

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyNext(VARIANT* xlIterator, VARIANT* xlElement, VARIANT_BOOL* xlResult)
{
	LogCtx log("PyNext");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlIterator));

		if(0 == PyIter_Check((PyObject*) obj))
			throw Exception() << "Object is not an iterator.";

		VariantInit(xlElement);
			
		PyObject* pRet = PyIter_Next(obj);
		if(NULL == pRet && NULL == PyErr_Occurred())
		{
			*xlResult = VARIANT_FALSE;
			return S_OK;
		}

		*xlResult = VARIANT_TRUE;

		PyNewRef ret(pRet);
		WrapPyToVariant(ret, xlElement);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall IPyObjectImpl::Next(ULONG celt, VARIANT* rgVar, ULONG* pCeltFetched)
{
	try
	{
		if(0 == PyIter_Check(this->pObj))
			throw Exception() << "Object is not an iterator.";

		int nFetched = 0;

		if (pCeltFetched != NULL)
			*pCeltFetched = 0;
		if (rgVar == NULL)
			return E_INVALIDARG;

		for(size_t k=0; k<celt; k++)
			VariantInit(&rgVar[k]);
			
		for(size_t k=0; k<celt; k++)
		{
			PyObject* pRet = PyIter_Next(this->pObj);
			if(NULL == pRet && NULL == PyErr_Occurred())
			{
				break;
			}
			else
			{
				PyNewRef ret(pRet);
				WrapPyToVariant(ret, &rgVar[k]);
				nFetched++;
				if (pCeltFetched != NULL)
					*pCeltFetched = nFetched;
			}
		}

		return nFetched == celt ? S_OK : S_FALSE;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyCall(VARIANT* xlObject, BSTR xlMethod, VARIANT* xlArgs, VARIANT* xlKeywordArgs, bool xlConsole, VARIANT* xlResult)
{
	LogCtx log("PyCall");

	try
	{
		ExecutionSentinel exe;

		if(xlConsole)
			ShowConsole(true);

		PyRef method;
		if(0 == SysStringLen(xlMethod))
			method = PyNewRef(VariantToPy(xlObject));
		else
		{
			std::string strMethod = BStrToStdString(xlMethod);
			
			PyNewRef obj(VariantToPy(xlObject));
			method = PyNewRef(PyObject_GetAttrString(obj, strMethod.c_str()));
		}

		PyRef args, kwargs;
		if(IsMissing(xlArgs))
			args = PyNewRef(PyTuple_New(0));
		else
		{
			args = PyNewRef(VariantToPy(xlArgs));
			if(!PyTuple_Check(args))
				throw Exception() << "Args should be a tuple object.";
		}
		if(!IsMissing(xlKeywordArgs))
		{
			kwargs = PyNewRef(VariantToPy(xlKeywordArgs));
			if(!PyDict_Check(kwargs))
				throw Exception() << "KwArgs should be a dict object.";
		}

		PyNewRef ret(PyObject_Call(method, args, kwargs));

		VariantInit(xlResult);
		WrapPyToVariant(ret, xlResult);

		if(xlConsole)
		{
			fprintf(stdout, "--- Press enter to close ---");
			fgetc(stdin);
			ShowConsole(false);
		}

		return S_OK;
	}
	catch(Exception& e)
	{
		if(xlConsole)
		{
			fprintf(stdout, "--- Error ---\n");
			fprintf(stdout, e.message().c_str());
			fprintf(stdout, "\n--- Press enter to close ---");
			fgetc(stdin);
			ShowConsole(false);
			return S_OK;
		}
		else
			return RaiseError(e.message());
	}
}

HRESULT __stdcall PyGetAttr(VARIANT* xlObject, BSTR xlAttribute, VARIANT* xlResult)
{
	LogCtx log("PyGetAttr");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlObject));

		std::string strAttr = BStrToStdString(xlAttribute);
		PyNewRef ret(PyObject_GetAttrString(obj, strAttr.c_str()));

		VariantInit(xlResult);
		WrapPyToVariant(ret, xlResult);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PySetAttr(VARIANT* xlObject, BSTR xlAttribute, VARIANT* xlValue)
{
	LogCtx log("PySetAttr");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlObject));
		PyNewRef val(VariantToPy(xlValue));

		std::string strAttr = BStrToStdString(xlAttribute);
		if(PyObject_SetAttrString(obj, strAttr.c_str(), val) < 0)
			throw PythonException();

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyGetItem(VARIANT* xlObject, VARIANT* xlKey, VARIANT* xlResult)
{
	LogCtx log("PyGetItem");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlObject));
		PyNewRef key(VariantToPy(xlKey));

		PyNewRef ret(PyObject_GetItem(obj, key));

		VariantInit(xlResult);
		WrapPyToVariant(ret, xlResult);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PySetItem(VARIANT* xlObject, VARIANT* xlKey, VARIANT* xlValue)
{
	LogCtx log("PySetItem");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlObject));
		PyNewRef key(VariantToPy(xlKey));
		PyNewRef val(VariantToPy(xlValue));

		if(PyObject_SetItem(obj, key, val) < 0)
			throw PythonException();

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyDelItem(VARIANT* xlObject, VARIANT* xlKey)
{
	LogCtx log("PyDelItem");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlObject));
		PyNewRef key(VariantToPy(xlKey));

		if(PyObject_DelItem(obj, key) < 0)
			throw PythonException();

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyContains(VARIANT* xlObject, VARIANT* xlKey, VARIANT_BOOL* xlResult)
{
	LogCtx log("PyContains");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlObject));
		PyNewRef key(VariantToPy(xlKey));

		int res = PySequence_Contains(obj, key);
		if(res < 0)
			throw PythonException();

		*xlResult = (res == 1) ? -1 : 0;
		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyDelAttr(VARIANT* xlObject, BSTR xlKey)
{
	LogCtx log("PyDelAttr");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlObject));
		std::string key = BStrToStdString(xlKey);

		if(PyObject_DelAttrString(obj, key.c_str()) < 0)
			throw PythonException();

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyHasAttr(VARIANT* xlObject, BSTR xlKey, VARIANT_BOOL* xlResult)
{
	LogCtx log("PyHasAttr");

	try
	{
		ExecutionSentinel exe;

		PyNewRef obj(VariantToPy(xlObject));
		std::string key = BStrToStdString(xlKey);

		int res = PyObject_HasAttrString(obj, key.c_str());
		if(res < 0)
			throw PythonException();

		*xlResult = (res == 1) ? -1 : 0;
		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

HRESULT __stdcall PyModule(BSTR xlModule, VARIANT_BOOL& reload, BSTR xlAddPath, BSTR xlPythonPath, VARIANT* xlResult)
{
	LogCtx log("PyModule");

	try
	{
		ExecutionSentinel exe;

		SetupPath(xlAddPath, xlPythonPath);

		std::string strModule = BStrToStdString(xlModule);
		PyNewRef module(PyImport_ImportModule(strModule.c_str()));
		if(reload != 0)
			module = PyNewRef(PyImport_ReloadModule(module));
		VariantInit(xlResult);
		WrapPyToVariant(module, xlResult);

		return S_OK;
	}
	catch(Exception& e)
	{
		return RaiseError(e.message());
	}
}

//long __stdcall PyCall(char*& xlPythonPath, const char*& xlModule, const char*& xlFunction, VARIANT* xlArgs, VARIANT* xlKeywordArgs, VARIANT* xlResult)
//{
//	try
//	{
//		PySys_SetPath(xlPythonPath);
//
//		PyNewRef module(PyImport_ImportModule(xlModule));
//
//		PyBorrowedRef function(PyDict_GetItemString(PyModule_GetDict(module), xlFunction), false);
//		if(function == NULL)
//			throw Exception() << "Function " << xlModule << "." << xlFunction << " not found";
//		
//		PyNewRef args(xlArgs->vt == VT_EMPTY ? PyTuple_New(0) : VariantToTuple(xlArgs));
//		PyNewRef kwargs(xlKeywordArgs->vt == VT_EMPTY ? NULL : VariantToDict(xlKeywordArgs), false);
//
//		PyNewRef ret(PyObject_Call(function, args, kwargs));
//
//		PyToVariant(ret, xlResult);
//	}
//	catch(Exception& e)
//	{
//		StrToVariant(e.message(), xlResult);
//		return -1;
//	}
//
//	return 0;
//}

//long __stdcall PyEvalAndCall(char*& xlPythonPath, const char*& xlExpression, VARIANT* xlArgs, VARIANT* xlKeywordArgs, VARIANT* xlResult)
//{
//	try
//	{
//		PySys_SetPath(xlPythonPath);
//
//		PyBorrowedRef module(PyImport_AddModule("__main__"));
//		PyBorrowedRef globals(PyModule_GetDict(module));
//
//		PyNewRef function(PyRun_String(xlExpression, Py_eval_input, globals, globals));
//		
//		PyNewRef args(xlArgs->vt == VT_EMPTY ? PyTuple_New(0) : VariantToTuple(xlArgs));
//		PyNewRef kwargs(xlKeywordArgs->vt == VT_EMPTY ? NULL : VariantToDict(xlKeywordArgs));
//
//		PyNewRef ret(PyObject_Call(function, args, kwargs));
//
//		VariantInit(xlResult);
//		PyToVariant(ret, xlResult);
//	}
//	catch(Exception& e)
//	{
//		StrToVariant(e.message(), xlResult);
//		return -1;
//	}
//
//	return 0;
//}
//
//HRESULT __stdcall PyShowConsole(BOOL* show)
//{
//	try
//	{
//		if(*show)
//		{
//			AllocConsole();
//			HWND hWnd = GetConsoleWindow();
//			EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
//			BringWindowToTop(hWnd);
//		}
//		else
//		{
//			FreeConsole();
//		}
//
//		return S_OK;
//	}
//	catch(Exception& e)
//	{
//		return RaiseError(e.message());
//	}
//}