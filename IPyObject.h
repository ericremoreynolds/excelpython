class __declspec(uuid("f9a3cfc8-4b97-4cec-b514-07761521eb2e"))
IPyObjectImpl : public IPyObj, public IEnumVARIANT
{
protected:
	IPyObjectImpl(PyObject* obj)
	{
		SetDLLGuid();
		pObj = obj;
		m_cRef = 1;
	}

	ULONG m_cRef;
	PyObject* pObj;
	GUID m_dllGuid;

	static GUID g_dllGuid;

	static void GenerateDLLGuid()
	{
		static bool bGenerated = false;
		if(!bGenerated)
		{
			CoCreateGuid(&g_dllGuid);
			bGenerated = true;
		}
	}

	void SetDLLGuid()
	{
		GenerateDLLGuid();
		CopyMemory(&m_dllGuid, &g_dllGuid, sizeof(GUID));
	}

	bool CheckDLLGuid()
	{
		GenerateDLLGuid();
		return TRUE == IsEqualGUID(m_dllGuid, g_dllGuid);
	}

public:
	static IPyObj* FromNewRef(PyObject* obj)
	{
		return new IPyObjectImpl(obj);
	}

	static IPyObj* FromBorrowedRef(PyObject* obj)
	{
		Py_XINCREF(obj);
		return new IPyObjectImpl(obj);
	}

	HRESULT __stdcall Str(VARIANT* Result);
	HRESULT __stdcall Var(LONG* Dimensions, VARIANT* Result);
	HRESULT __stdcall get_Item(VARIANT* Key, IPyObj** Result);
	HRESULT __stdcall get_Attr(BSTR Attribute, IPyObj** Result);
	HRESULT __stdcall get_Count(LONG *pVal);
	HRESULT __stdcall get__NewEnum(IUnknown** ppUnk);

	PyObject* GetNewRef()
	{
		if(!CheckDLLGuid())
			throw Exception() << "Alien COM reference to Python object: this probably means ExcelPython DLL has been reloaded.";
		Py_XINCREF(pObj);
		return pObj;
	}

	PyObject* GetBorrowedRef()
	{
		if(!CheckDLLGuid())
			throw Exception() << "Alien COM reference to Python object: this probably means ExcelPython DLL has been reloaded.";
		return pObj;
	}

	~IPyObjectImpl()
	{
		if(CheckDLLGuid())
			Py_XDECREF(pObj);
	}

	HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObj)
	{
		if(!ppvObj)
			return E_INVALIDARG;

		*ppvObj = NULL;
		if (riid == IID_IUnknown)
		{
			*ppvObj = (LPVOID) (IUnknown*) (IPyObj*) this;
			AddRef();
			return S_OK;
		}
		if (riid == __uuidof(IPyObjectImpl))
		{
			*ppvObj = (LPVOID) this;
			AddRef();
			return S_OK;
		}
		if (riid == __uuidof(IPyObj))
		{
			*ppvObj = (LPVOID) (IPyObj*) this;
			AddRef();
			return S_OK;
		}
		if (riid == __uuidof(IDispatch))
		{
			*ppvObj = (LPVOID) (IDispatch*) this;
			AddRef();
			return S_OK;
		}
		if (riid == __uuidof(IEnumVARIANT))
		{
			*ppvObj = (LPVOID) (IEnumVARIANT*) this;
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	ULONG __stdcall AddRef()
	{
		return ++m_cRef;
	}

	ULONG __stdcall Release()
	{
		ULONG ulRefCount = --m_cRef;
		if (0 == m_cRef)
			delete this;

		return ulRefCount;
	}

	HRESULT __stdcall GetIDsOfNames(
		REFIID riid,
		LPOLESTR *rgszNames,
		UINT cNames,
		LCID lcid,
		DISPID *rgDispId
	)
	{
		ITypeInfo* pTypeInfo;
		HRESULT hr = LoadTypeInfo(&pTypeInfo);
		if(FAILED(hr))
			return hr;

		return DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
	}

	static HRESULT LoadTypeInfo(ITypeInfo** ppTypeInfo)
	{
		*ppTypeInfo = NULL;
		static ITypeInfo* pTypeInfo = NULL;
		if(pTypeInfo == NULL)
		{
			HRESULT hr;

			// Load the type library.
			ITypeLib* pTypeLib;
			if(FAILED(hr = LoadRegTypeLib(LIBID_ExcelPython, 1, 0, LANG_NEUTRAL, &pTypeLib)))
				return hr;

			// Get the type info.
			if(FAILED(hr = pTypeLib->GetTypeInfoOfGuid(IID_IPyObj, &pTypeInfo)))
				return hr;

			// Release the type library.
			pTypeLib->Release();
		}
		*ppTypeInfo = pTypeInfo;
		return S_OK;
	}

	HRESULT __stdcall GetTypeInfo(
		UINT iTInfo,
		LCID lcid,
		ITypeInfo **ppTInfo
	)
	{
		*ppTInfo = NULL;

		if(iTInfo != 0)
			return DISP_E_BADINDEX;

		HRESULT hr = LoadTypeInfo(ppTInfo);
		if(FAILED(hr))
			return hr;

		(*ppTInfo)->AddRef();
		return S_OK;
	}

	HRESULT __stdcall GetTypeInfoCount(
		UINT *pctinfo
	)
	{
		*pctinfo = 1;
		return S_OK;
	}

	HRESULT __stdcall Invoke(
		DISPID dispIdMember,
		REFIID riid,
		LCID lcid,
		WORD wFlags,
		DISPPARAMS *pDispParams,
		VARIANT *pVarResult,
		EXCEPINFO *pExcepInfo,
		UINT *puArgErr
	)
	{
		ITypeInfo* pTypeInfo;
		HRESULT hr = LoadTypeInfo(&pTypeInfo);
		if(FAILED(hr))
			return hr;

		return DispInvoke(this, pTypeInfo, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	}

    // Return the next 0 or more items.
    HRESULT __stdcall Next(ULONG celt, VARIANT* rgVar, ULONG* pCeltFetched);

    // Skip the next 0 or more items.
    HRESULT __stdcall Skip(ULONG celt)
	{
		return E_NOTIMPL;
	}

    // Reset the enumerator.
    HRESULT __stdcall Reset()
	{
		return E_NOTIMPL;
	}

    // Create a new enumerator with the identical state.
    HRESULT __stdcall Clone(IEnumVARIANT ** ppEnum)
	{
		*ppEnum = NULL;
		return E_NOTIMPL;
	}
};

_COM_SMARTPTR_TYPEDEF(IPyObjectImpl, __uuidof(IPyObjectImpl));