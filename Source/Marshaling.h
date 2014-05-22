
void WrapPyToVariant(PyObject* obj, VARIANT* var);
void WrapPyToIPyObject(PyObject* obj, IPyObj** pPy);
void ConvertPyToVariant(PyObject* obj, VARIANT* var, bool deep = false, int dims = -1);
void StrToVariant(const std::string& str, VARIANT* var);
std::string WCharToStdString(const wchar_t* str);
std::string BStrToStdString(BSTR str);
BSTR CStrToBStr(const char* str);

// return new references
PyObject* BStrToPyString(BSTR str);
PyObject* VariantToPy(VARIANT* var, int dims = -1);
PyObject* VariantToTuple(VARIANT* var);
int SafeArrayLength(SAFEARRAY* pSA);
PyObject* SafeArrayToTuple(SAFEARRAY* pSA);
PyObject* SafeArrayToList(SAFEARRAY* pSA);
PyObject* VariantToDict(VARIANT* var);
PyObject* VariantToList(VARIANT* var, int dims);
PyObject* SafeArrayToDict(SAFEARRAY* pSA);

class AutoSafeArrayAccessData
{
	SAFEARRAY* _pSA;

public:
	AutoSafeArrayAccessData(SAFEARRAY* pSA, void** ppData)
	{
		_pSA = NULL;
		if(FAILED(SafeArrayAccessData(pSA, ppData)))
			throw Exception() << "Could not access safe array data.";
		_pSA = pSA;
	}

	~AutoSafeArrayAccessData()
	{
		if(_pSA != NULL && FAILED(SafeArrayUnaccessData(_pSA)))
			throw Exception() << "Could not unaccess safe array data.";
	}
};

class AutoSafeArrayCreate
{
public:
	SAFEARRAY* pSafeArray;

	AutoSafeArrayCreate(VARTYPE vt, UINT cDims, SAFEARRAYBOUND* rgsabound)
	{
		pSafeArray = SafeArrayCreate(vt, cDims, rgsabound);
		if(pSafeArray == NULL)
			throw Exception() << "Could not create safe array.";
	}

	~AutoSafeArrayCreate()
	{
		if(pSafeArray != NULL && FAILED(SafeArrayDestroy(pSafeArray)))
			throw Exception() << "Could not destroy safe array.";
	}
};