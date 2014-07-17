#include <boost/scoped_array.hpp>

#include <windows.h>
#include <comdef.h>
#include <fstream>
#include <tchar.h>
#include "Logging.h"
#include "Python.h"
#include "Exception.h"

#include "ExcelPython_h.h"

#include "ConsoleObject.h"

extern HINSTANCE hModuleDLL;

class PyRef
{
protected:
	PyRef(PyObject* obj, bool throwIfNull = true)
	{
		if(throwIfNull && obj == NULL)
			throw PythonException();

		ptr = obj;
	}

public:
	PyObject* ptr;

	PyRef(const PyRef& ref)
	{
		ptr = ref.ptr;
		Py_XINCREF(ptr);
	}

	PyRef()
	{
		ptr = NULL;
	}

	~PyRef()
	{
		Py_XDECREF(ptr);
	}

	bool operator==(PyObject* obj)
	{
		return ptr == obj;
	}

	bool operator!=(PyObject* obj)
	{
		return ptr != obj;
	}

	operator PyObject* ()
	{
		return ptr;
	}

	PyObject* detach()
	{
		PyObject* ret = ptr;
		ptr = NULL;
		return ret;
	}

	PyRef& operator= (const PyRef& ref)
	{
		if(ptr != NULL)
			Py_XDECREF(ptr);
		ptr = ref.ptr;
		Py_XINCREF(ptr);
		return *this;
	}
};

class PyNewRef : public PyRef
{
public:
	PyNewRef(PyObject* obj, bool throwIfNull = true)
		: PyRef(obj, throwIfNull)
	{
		//Py_XINCREF(obj);
	}
};

class PyBorrowedRef : public PyRef
{
public:
	PyBorrowedRef(PyObject* obj, bool throwIfNull = true)
		: PyRef(obj, throwIfNull)
	{
		Py_XINCREF(obj);
	}
};

#include "Marshaling.h"
#include "PyDispatch.h"
#include "IPyObject.h"